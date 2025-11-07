import { GraphService } from './GraphService';
import { ListService } from './ListService';
import { cacheHelper } from '../utils/CacheHelper';
import { IUserProfile, ISearchResult, ISearchFilter } from '../models/IUserProfile';
import { Constants } from '../models/Constants';

/**
 * Orchestration service that manages user data across multiple caches
 * Implements multi-tier caching strategy:
 * 1. IndexedDB (client-side, 30 min TTL)
 * 2. SharePoint List (tenant-wide, 24 hour TTL)
 * 3. Microsoft Graph (source of truth)
 */
export class PeopleService {
  private graphService: GraphService;
  private listService: ListService;

  constructor(graphService: GraphService, listService: ListService) {
    this.graphService = graphService;
    this.listService = listService;
  }

  /**
   * Initialize services
   */
  public async initialize(): Promise<void> {
    await Promise.all([
      cacheHelper.initialize(),
      this.listService.ensureList()
    ]);
  }

  /**
   * Search for users with multi-tier caching
   */
  public async searchUsers(filter: ISearchFilter, pageSize: number = Constants.DEFAULT_PAGE_SIZE): Promise<ISearchResult> {
    const cacheKey = this.getCacheKey('search', filter);

    try {
      // Try client-side cache first
      const cachedResult = await cacheHelper.get<ISearchResult>(cacheKey);
      if (cachedResult) {
        console.log('Returning cached search results');
        return cachedResult;
      }

      // Search from Graph (source of truth)
      const result = await this.graphService.searchUsers(filter.searchText, pageSize);

      // Enrich with photos asynchronously (don't block)
      this.enrichUsersWithPhotos(result.users).catch(console.error);

      // Cache the result
      await cacheHelper.set(cacheKey, result);

      // Update list cache asynchronously for top results
      this.updateListCache(result.users.slice(0, 20)).catch(console.error);

      return result;
    } catch (error) {
      console.error('Error searching users:', error);
      // Fallback to list cache if Graph fails
      try {
        const listUsers = await this.listService.searchUsers(filter.searchText, pageSize);
        return {
          users: listUsers,
          totalCount: listUsers.length,
          hasMore: false
        };
      } catch (listError) {
        console.error('Error falling back to list cache:', listError);
        throw error;
      }
    }
  }

  /**
   * Get user by ID with real-time verification
   * Always checks Graph API to ensure data is current
   */
  public async getUserById(userId: string, forceRefresh = false): Promise<IUserProfile | null> {
    const cacheKey = this.getCacheKey('user', userId);

    try {
      // If not forcing refresh, try client cache first
      if (!forceRefresh) {
        const cachedUser = await cacheHelper.get<IUserProfile>(cacheKey);
        if (cachedUser) {
          // Return cached user but verify in background
          this.verifyAndUpdateUser(userId).catch(console.error);
          return cachedUser;
        }
      }

      // Get fresh data from Graph
      const user = await this.graphService.getUserById(userId);
      if (!user) return null;

      // Get photo
      const photoUrl = await this.graphService.getUserPhoto(userId);
      if (photoUrl) {
        user.photoUrl = photoUrl;
      }

      // Cache the user
      await cacheHelper.set(cacheKey, user);

      // Update list cache asynchronously
      this.listService.upsertUser(user).catch(console.error);

      return user;
    } catch (error) {
      console.error('Error getting user by ID:', error);
      // Fallback to list cache
      try {
        return await this.listService.getUser(userId);
      } catch (listError) {
        console.error('Error falling back to list cache:', listError);
        throw error;
      }
    }
  }

  /**
   * Get user by User Principal Name
   */
  public async getUserByUPN(upn: string, forceRefresh = false): Promise<IUserProfile | null> {
    return this.getUserById(upn, forceRefresh);
  }

  /**
   * Get all departments (with caching)
   */
  public async getDepartments(): Promise<string[]> {
    const cacheKey = this.getCacheKey('departments', 'all');

    try {
      // Try cache first
      const cached = await cacheHelper.get<string[]>(cacheKey);
      if (cached) {
        return cached;
      }

      // Try list cache (faster than Graph for large orgs)
      const listDepartments = await this.listService.getDepartments();
      if (listDepartments.length > 0) {
        await cacheHelper.set(cacheKey, listDepartments, Constants.LIST_CACHE_TTL);
        return listDepartments;
      }

      // Fallback to Graph (expensive for large orgs)
      const graphDepartments = await this.graphService.getDepartments();
      await cacheHelper.set(cacheKey, graphDepartments, Constants.LIST_CACHE_TTL);

      return graphDepartments;
    } catch (error) {
      console.error('Error getting departments:', error);
      return [];
    }
  }

  /**
   * Get all office locations (with caching)
   */
  public async getOfficeLocations(): Promise<string[]> {
    const cacheKey = this.getCacheKey('locations', 'all');

    try {
      // Try cache first
      const cached = await cacheHelper.get<string[]>(cacheKey);
      if (cached) {
        return cached;
      }

      // Get from Graph
      const locations = await this.graphService.getOfficeLocations();
      await cacheHelper.set(cacheKey, locations, Constants.LIST_CACHE_TTL);

      return locations;
    } catch (error) {
      console.error('Error getting office locations:', error);
      return [];
    }
  }

  /**
   * Clear all caches
   */
  public async clearAllCaches(): Promise<void> {
    await Promise.all([
      cacheHelper.clear(),
      this.listService.clearCache()
    ]);
  }

  /**
   * Verify user data against Graph and update if changed
   */
  private async verifyAndUpdateUser(userId: string): Promise<void> {
    try {
      const freshUser = await this.graphService.getUserById(userId);
      if (!freshUser) return;

      const photoUrl = await this.graphService.getUserPhoto(userId);
      if (photoUrl) {
        freshUser.photoUrl = photoUrl;
      }

      // Update caches
      const cacheKey = this.getCacheKey('user', userId);
      await cacheHelper.set(cacheKey, freshUser);
      await this.listService.upsertUser(freshUser);
    } catch (error) {
      console.error('Error verifying user:', error);
    }
  }

  /**
   * Enrich users with profile photos
   */
  private async enrichUsersWithPhotos(users: IUserProfile[]): Promise<void> {
    const photoPromises = users.map(async (user) => {
      try {
        const photoUrl = await this.graphService.getUserPhoto(user.id);
        if (photoUrl) {
          user.photoUrl = photoUrl;
        }
      } catch (error) {
        // Photo fetch failures are not critical
        console.debug(`Failed to fetch photo for ${user.displayName}`);
      }
    });

    await Promise.allSettled(photoPromises);
  }

  /**
   * Update list cache with user data
   */
  private async updateListCache(users: IUserProfile[]): Promise<void> {
    const updatePromises = users.map(user => this.listService.upsertUser(user));
    await Promise.allSettled(updatePromises);
  }

  /**
   * Generate cache key
   */
  private getCacheKey(type: string, identifier: string | ISearchFilter): string {
    if (typeof identifier === 'string') {
      return `${type}:${identifier}`;
    }

    // For search filters, create a deterministic key
    const filter = identifier as ISearchFilter;
    return `${type}:${filter.searchText}:${filter.department || ''}:${filter.officeLocation || ''}:${filter.jobTitle || ''}`;
  }

  /**
   * Get cache statistics
   */
  public async getCacheStats(): Promise<{ clientCache: number; listCache: number }> {
    try {
      const clientStats = await cacheHelper.getStats();
      // List cache count would require a separate query
      return {
        clientCache: clientStats.count,
        listCache: 0 // TODO: Implement list cache count
      };
    } catch (error) {
      console.error('Error getting cache stats:', error);
      return { clientCache: 0, listCache: 0 };
    }
  }
}
