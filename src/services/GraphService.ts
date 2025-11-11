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
 * Service for interacting with Microsoft Graph API
 */
export class GraphService {
  private graphClient: MSGraphClientV3;

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  /**
   * Search users by query string
   * Uses $search query parameter for server-side filtering
   */
  public async searchUsers(
    searchText: string,
    pageSize: number = Constants.DEFAULT_PAGE_SIZE,
    skipToken?: string
  ): Promise<ISearchResult> {
    try {
      let endpoint = `/users?$select=${Constants.GRAPH_SELECT_FIELDS}&$top=${pageSize}&$count=true`;

      // Build filter query
      if (searchText && searchText.length >= Constants.MIN_SEARCH_LENGTH) {
        // Use $filter for more precise matching
        const filter = `startswith(displayName,'${this.escapeODataString(searchText)}') or ` +
                      `startswith(mail,'${this.escapeODataString(searchText)}') or ` +
                      `startswith(surname,'${this.escapeODataString(searchText)}') or ` +
                      `startswith(givenName,'${this.escapeODataString(searchText)}') or ` +
                      `startswith(department,'${this.escapeODataString(searchText)}')`;
        endpoint += `&$filter=${filter}`;
      }

      // Add pagination token if provided
      if (skipToken) {
        endpoint += `&$skiptoken=${skipToken}`;
      }

      const response = await this.graphClient
        .api(endpoint)
        .header('ConsistencyLevel', 'eventual')
        .get();

      const users: IUserProfile[] = response.value.map((user: any) => this.mapGraphUserToProfile(user));

      return {
        users,
        totalCount: response['@odata.count'] || users.length,
        hasMore: !!response['@odata.nextLink'],
        nextLink: response['@odata.nextLink']
      };
    } catch (error) {
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
      if ((error as any).statusCode === 404) {
        return null;
      }
      throw new Error(ErrorHandler.getUserMessage(error, 'GraphService.getUserById'));
    }
  }

  /**
   * Get user profile photo
   */
  public async getUserPhoto(userId: string): Promise<string | null> {
    try {
      const photoBlob = await this.graphClient
        .api(`/users/${userId}/photos/${Constants.PHOTO_SIZE}/$value`)
        .get();

      // Convert blob to base64 data URL
      return await this.blobToDataURL(photoBlob);
    } catch (error) {
      // Photo not found is not an error condition
      if ((error as any).statusCode === 404) {
        return null;
      }
      console.warn('Failed to fetch user photo:', error);
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
      if ((error as any).statusCode === 404) {
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
        const endpoint = nextLink || '/users?$select=id,department&$top=999';
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
        const endpoint = nextLink || '/users?$select=id,officeLocation&$top=999';
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

        const batchResponse = await this.graphClient
          .api('/$batch')
          .post({ requests: batchRequests });

        batchResponse.responses.forEach((response: any) => {
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
   * Map Graph API user object to IUserProfile
   */
  private mapGraphUserToProfile(graphUser: any): IUserProfile {
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
}
