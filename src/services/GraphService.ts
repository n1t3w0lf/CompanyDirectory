import { MSGraphClientV3 } from '@microsoft/sp-http';
import { IUserProfile, ISearchResult } from '../models/IUserProfile';
import { Constants } from '../models/Constants';
import { ErrorHandler } from '../utils/ErrorHandler';

/**
 * Graph API Response Types
 */
interface IGraphResponse {
  value: IGraphUser[];
  '@odata.count'?: number;
  '@odata.nextLink'?: string;
}

interface IGraphUser {
  id: string;
  userPrincipalName: string;
  displayName: string;
  givenName: string;
  surname: string;
  mail: string;
  jobTitle: string;
  department: string;
  officeLocation: string;
  businessPhones: string[];
  mobilePhone: string;
  city: string;
  country: string;
  companyName: string;
  preferredLanguage: string;
  employeeId: string;
}

/**
 * Microsoft Graph API Error Response
 */
interface IGraphError {
  statusCode?: number;
  code?: string;
  message?: string;
}

/**
 * Microsoft Graph Batch API Response
 */
interface IGraphBatchResponse {
  responses: Array<{
    id: string;
    status: number;
    body: IGraphUser;
  }>;
}

/**
 * Service for interacting with Microsoft Graph API
 */
export class GraphService {
  private graphClient: MSGraphClientV3;

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  /**
   * Search users by query string using the Graph $search parameter (tokenized,
   * order-independent partial matching on displayName; prefix match on other fields).
   * @param nextLinkUrl - Full nextLink URL from previous response for pagination
   */
  public async searchUsers(
    searchText: string,
    pageSize: number = Constants.DEFAULT_PAGE_SIZE,
    nextLinkUrl?: string
  ): Promise<ISearchResult> {
    try {
      let apiRequest;

      // If we have a nextLink URL, use it directly for pagination
      if (nextLinkUrl) {
        // Extract the path and query from the full nextLink URL
        // nextLink format: https://graph.microsoft.com/v1.0/users?...&$skiptoken=...
        const endpoint = this.parseNextLink(nextLinkUrl);
        apiRequest = this.graphClient.api(endpoint);
      } else {
        // Build initial request
        let endpoint = `/users?$select=${Constants.GRAPH_SELECT_FIELDS}&$top=${pageSize}&$count=true`;

        // Build $search query (tokenized partial matching). displayName is
        // tokenized (matches mid-name terms in any order); other fields fall
        // back to prefix matching. Requires ConsistencyLevel: eventual (below).
        if (searchText && searchText.length >= Constants.MIN_SEARCH_LENGTH) {
          const term = this.escapeSearchTerm(searchText);
          const searchClause =
            `"displayName:${term}" OR "mail:${term}" OR ` +
            `"givenName:${term}" OR "surname:${term}" OR "department:${term}"`;
          // The Graph client assembles the query string verbatim (no encoding of
          // its own), so encode the $search value ourselves.
          endpoint += `&$search=${encodeURIComponent(searchClause)}`;
        }

        apiRequest = this.graphClient.api(endpoint);
      }

      const response = await apiRequest
        .header('ConsistencyLevel', 'eventual')
        .get();

      const users: IUserProfile[] = response.value.map((user: IGraphUser) => this.mapGraphUserToProfile(user));

      return {
        users,
        totalCount: response['@odata.count'] || users.length,
        hasMore: !!response['@odata.nextLink'],
        nextLink: response['@odata.nextLink']
      };
    } catch (error) {
      console.error('GraphService.searchUsers error details:', {
        error,
        nextLinkUrl,
        searchText,
        pageSize
      });
      throw new Error(ErrorHandler.getUserMessage(error, 'GraphService.searchUsers'));
    }
  }

  /**
   * Get user by ID or UPN
   */
  public async getUserById(userId: string): Promise<IUserProfile | null> {
    try {
      const user = await this.graphClient
        .api(`/users/${userId}`)
        .select(Constants.GRAPH_SELECT_FIELDS)
        .get();

      return this.mapGraphUserToProfile(user);
    } catch (error) {
      if ((error as IGraphError).statusCode === 404) {
        return null;
      }
      throw new Error(ErrorHandler.getUserMessage(error, 'GraphService.getUserById'));
    }
  }

  /**
   * Get user profile photo.
   * @param throwOnError - when true, re-throws non-404 errors (e.g. 429 throttling)
   *   so callers can distinguish a transient failure from a genuine "no photo"
   *   (404 -> null). Defaults to false to preserve existing callers' behaviour.
   */
  public async getUserPhoto(userId: string, throwOnError = false): Promise<string | null> {
    try {
      const photoBlob = await this.graphClient
        .api(`/users/${userId}/photos/${Constants.PHOTO_SIZE}/$value`)
        .get();

      // Convert blob to base64 data URL
      return await this.blobToDataURL(photoBlob);
    } catch (error) {
      // Photo not found (404) is a genuine "no photo", not an error condition.
      if ((error as IGraphError).statusCode === 404) {
        return null;
      }
      console.warn('Failed to fetch user photo:', error);
      if (throwOnError) {
        throw error;
      }
      return null;
    }
  }

  /**
   * Get user's manager
   */
  public async getUserManager(userId: string): Promise<IUserProfile | null> {
    try {
      const manager = await this.graphClient
        .api(`/users/${userId}/manager`)
        .select(Constants.GRAPH_SELECT_FIELDS)
        .get();

      return this.mapGraphUserToProfile(manager);
    } catch (error) {
      if ((error as IGraphError).statusCode === 404) {
        return null;
      }
      console.warn('Failed to fetch user manager:', error);
      return null;
    }
  }

  /**
   * Get all departments (aggregated from users)
   * Note: This is expensive for large orgs. Consider caching or using a separate API
   * Updated to paginate through ALL users to support 60K+ organizations
   */
  public async getDepartments(): Promise<string[]> {
    try {
      const departments = new Set<string>();
      let nextLink: string | undefined;

      // Paginate through all users to get complete department list
      do {
        // Use nextLink if available, otherwise use initial endpoint
        const endpoint = nextLink ? this.parseNextLink(nextLink) : '/users?$select=id,department&$top=999';
        const response: IGraphResponse = await this.graphClient
          .api(endpoint)
          .header('ConsistencyLevel', 'eventual')
          .get();

        response.value.forEach((user: IGraphUser) => {
          if (user.department) {
            departments.add(user.department);
          }
        });

        nextLink = response['@odata.nextLink'];
      } while (nextLink);

      return Array.from(departments).sort();
    } catch (error) {
      throw new Error(ErrorHandler.getUserMessage(error, 'GraphService.getDepartments'));
    }
  }

  /**
   * Get all office locations
   * Updated to paginate through ALL users to support 60K+ organizations
   */
  public async getOfficeLocations(): Promise<string[]> {
    try {
      const locations = new Set<string>();
      let nextLink: string | undefined;

      // Paginate through all users to get complete location list
      do {
        // Use nextLink if available, otherwise use initial endpoint
        const endpoint = nextLink ? this.parseNextLink(nextLink) : '/users?$select=id,officeLocation&$top=999';
        const response: IGraphResponse = await this.graphClient
          .api(endpoint)
          .header('ConsistencyLevel', 'eventual')
          .get();

        response.value.forEach((user: IGraphUser) => {
          if (user.officeLocation) {
            locations.add(user.officeLocation);
          }
        });

        nextLink = response['@odata.nextLink'];
      } while (nextLink);

      return Array.from(locations).sort();
    } catch (error) {
      throw new Error(ErrorHandler.getUserMessage(error, 'GraphService.getOfficeLocations'));
    }
  }

  /**
   * Batch get multiple users by IDs
   */
  public async getUsersBatch(userIds: string[]): Promise<IUserProfile[]> {
    try {
      const batchSize = Constants.GRAPH_BATCH_SIZE;
      const batches: string[][] = [];

      // Split into batches
      for (let i = 0; i < userIds.length; i += batchSize) {
        batches.push(userIds.slice(i, i + batchSize));
      }

      // Execute batches sequentially to avoid rate limiting
      const allUsers: IUserProfile[] = [];
      for (const batch of batches) {
        const batchRequests = batch.map((id, index) => ({
          id: index.toString(),
          method: 'GET',
          url: `/users/${id}?$select=${Constants.GRAPH_SELECT_FIELDS}`
        }));

        const batchResponse: IGraphBatchResponse = await this.graphClient
          .api('/$batch')
          .post({ requests: batchRequests });

        batchResponse.responses.forEach((response) => {
          if (response.status === 200) {
            allUsers.push(this.mapGraphUserToProfile(response.body));
          }
        });
      }

      return allUsers;
    } catch (error) {
      throw new Error(ErrorHandler.getUserMessage(error, 'GraphService.getUsersBatch'));
    }
  }

  /**
   * Extract API path from nextLink URL
   * Removes the version prefix (/v1.0 or /beta) that MSGraphClientV3 adds automatically
   */
  private parseNextLink(nextLinkUrl: string): string {
    const url = new URL(nextLinkUrl);
    let pathname = url.pathname;

    // Remove version prefix since MSGraphClientV3 adds it automatically
    if (pathname.startsWith('/v1.0/')) {
      pathname = pathname.substring(5); // Remove '/v1.0'
    } else if (pathname.startsWith('/beta/')) {
      pathname = pathname.substring(6); // Remove '/beta'
    }

    return pathname + url.search;
  }

  /**
   * Map Graph API user object to IUserProfile
   */
  private mapGraphUserToProfile(graphUser: IGraphUser): IUserProfile {
    return {
      id: graphUser.id,
      userPrincipalName: graphUser.userPrincipalName,
      displayName: graphUser.displayName,
      givenName: graphUser.givenName,
      surname: graphUser.surname,
      mail: graphUser.mail,
      jobTitle: graphUser.jobTitle,
      department: graphUser.department,
      officeLocation: graphUser.officeLocation,
      businessPhones: graphUser.businessPhones || [],
      mobilePhone: graphUser.mobilePhone,
      city: graphUser.city,
      country: graphUser.country,
      companyName: graphUser.companyName,
      preferredLanguage: graphUser.preferredLanguage,
      employeeId: graphUser.employeeId,
      lastVerified: new Date()
    };
  }

  /**
   * Convert Blob to Data URL
   */
  private async blobToDataURL(blob: Blob): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result as string);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  /**
   * Escape OData string for filter queries
   */
  private escapeODataString(str: string): string {
    return str.replace(/'/g, "''");
  }

  /**
   * Sanitise a term for use inside a KQL $search clause. Double quotes delimit
   * clauses and backslash is an escape char, so both are stripped to prevent a
   * term from breaking out of its "property:term" clause. Ampersands are
   * replaced with a space: once URL-encoded they become %26, which trips a
   * documented Graph v1.0 bug that 400s $search on directory objects. Replacing
   * (not stripping) preserves token matching, e.g. "R&D" -> "R D" still matches.
   * The value is URL-encoded by the caller.
   */
  private escapeSearchTerm(str: string): string {
    return str.replace(/["\\]/g, '').replace(/&/g, ' ').replace(/\s+/g, ' ').trim();
  }
}
