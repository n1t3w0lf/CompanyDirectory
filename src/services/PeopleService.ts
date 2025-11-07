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
      // Try cache first (but only if it has data)
      const cached = await cacheHelper.get<string[]>(cacheKey);
      if (cached && cached.length > 0) {
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

      // Only cache if we have data
      if (graphDepartments.length > 0) {
        await cacheHelper.set(cacheKey, graphDepartments, Constants.LIST_CACHE_TTL);
      }

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
      // Try cache first (but only if it has data)
      const cached = await cacheHelper.get<string[]>(cacheKey);
      if (cached && cached.length > 0) {
        return cached;
      }

      // Try list cache (faster than Graph for large orgs)
      const listLocations = await this.listService.getOfficeLocations();
      if (listLocations.length > 0) {
        await cacheHelper.set(cacheKey, listLocations, Constants.LIST_CACHE_TTL);
        return listLocations;
      }

      // Fallback to Graph
      const locations = await this.graphService.getOfficeLocations();

      // Only cache if we have data
      if (locations.length > 0) {
        await cacheHelper.set(cacheKey, locations, Constants.LIST_CACHE_TTL);
      }

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
   * Get initial users for display (first 30 from list)
   */
  public async getInitialUsers(pageSize = 30): Promise<IUserProfile[]> {
    try {
      // Get users from SharePoint list (primary data source now)
      const users = await this.listService.getPaginatedUsers(pageSize, 1, 'Title');

      // Enrich with photos asynchronously (don't block)
      this.enrichUsersWithPhotos(users).catch(console.error);

      return users;
    } catch (error) {
      console.error('Error getting initial users:', error);
      return [];
    }
  }

  /**
   * Get filtered users with advanced filters
   */
  public async getFilteredUsers(filters: {
    searchText?: string;
    department?: string;
    officeLocation?: string;
    city?: string;
    country?: string;
    jobTitle?: string;
  }, pageSize = 50): Promise<IUserProfile[]> {
    try {
      let users: IUserProfile[] = [];

      // If there's a search text, search in the list
      if (filters.searchText && filters.searchText.length >= Constants.MIN_SEARCH_LENGTH) {
        users = await this.listService.searchUsers(filters.searchText, pageSize);
      } else {
        // Get filtered users from list
        users = await this.listService.getFilteredUsers({
          department: filters.department,
          officeLocation: filters.officeLocation,
          city: filters.city,
          country: filters.country,
          jobTitle: filters.jobTitle
        }, pageSize);
      }

      // Apply additional client-side filtering if needed
      if (filters.searchText && users.length > 0) {
        const searchLower = filters.searchText.toLowerCase();
        users = users.filter(u =>
          u.displayName?.toLowerCase().includes(searchLower) ||
          u.mail?.toLowerCase().includes(searchLower) ||
          u.department?.toLowerCase().includes(searchLower) ||
          u.jobTitle?.toLowerCase().includes(searchLower)
        );
      }

      return users;
    } catch (error) {
      console.error('Error getting filtered users:', error);
      return [];
    }
  }

  /**
   * Search Active Directory (Microsoft Graph) directly
   * Used as fallback when list search returns no results
   */
  public async searchActiveDirectory(searchText: string, pageSize = 50): Promise<IUserProfile[]> {
    try {
      if (!searchText || searchText.length < Constants.MIN_SEARCH_LENGTH) {
        return [];
      }

      console.log('Searching Active Directory for:', searchText);

      // Search Graph directly without caching
      const result = await this.graphService.searchUsers(searchText, pageSize);

      // Enrich with photos asynchronously
      this.enrichUsersWithPhotos(result.users).catch(err =>
        console.warn('Failed to enrich AD search results with photos:', err)
      );

      return result.users;
    } catch (error) {
      console.error('Error searching Active Directory:', error);
      return [];
    }
  }

  /**
   * Get all unique cities from list
   */
  public async getCities(): Promise<string[]> {
    const cacheKey = this.getCacheKey('cities', 'all');

    try {
      // Try cache first (but only if it has data)
      const cached = await cacheHelper.get<string[]>(cacheKey);
      if (cached && cached.length > 0) {
        return cached;
      }

      // Get from list cache
      const cities = await this.listService.getCities();

      // Only cache if we have data
      if (cities.length > 0) {
        await cacheHelper.set(cacheKey, cities, Constants.LIST_CACHE_TTL);
      }

      return cities;
    } catch (error) {
      console.error('Error getting cities:', error);
      return [];
    }
  }

  /**
   * Get all unique countries from list
   */
  public async getCountries(): Promise<string[]> {
    const cacheKey = this.getCacheKey('countries', 'all');

    try {
      // Try cache first
      const cached = await cacheHelper.get<string[]>(cacheKey);
      if (cached) {
        return cached;
      }

      // This would require a custom implementation or aggregation
      // For now, return empty array - can be enhanced later
      const countries: string[] = [];
      await cacheHelper.set(cacheKey, countries, Constants.LIST_CACHE_TTL);

      return countries;
    } catch (error) {
      console.error('Error getting countries:', error);
      return [];
    }
  }

  /**
   * Get all unique job titles from list
   */
  public async getJobTitles(): Promise<string[]> {
    const cacheKey = this.getCacheKey('jobtitles', 'all');

    try {
      // Try cache first
      const cached = await cacheHelper.get<string[]>(cacheKey);
      if (cached) {
        return cached;
      }

      // This would require a custom implementation or aggregation
      // For now, return empty array - can be enhanced later
      const jobTitles: string[] = [];
      await cacheHelper.set(cacheKey, jobTitles, Constants.LIST_CACHE_TTL);

      return jobTitles;
    } catch (error) {
      console.error('Error getting job titles:', error);
      return [];
    }
  }

  /**
   * Get total user count
   */
  public async getTotalUserCount(): Promise<number> {
    try {
      return await this.listService.getTotalUserCount();
    } catch (error) {
      console.error('Error getting total user count:', error);
      return 0;
    }
  }

  /**
   * Get cache statistics
   */
  public async getCacheStats(): Promise<{ clientCache: number; listCache: number }> {
    try {
      const clientStats = await cacheHelper.getStats();
      const listCount = await this.listService.getTotalUserCount();
      return {
        clientCache: clientStats.count,
        listCache: listCount
      };
    } catch (error) {
      console.error('Error getting cache stats:', error);
      return { clientCache: 0, listCache: 0 };
    }
  }
}
