import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import { IUserProfile } from '../models/IUserProfile';
import { Constants } from '../models/Constants';
import { ErrorHandler } from '../utils/ErrorHandler';

/**
 * SharePoint list item structure for cached user data
 * Note: All custom fields use PD_ prefix to avoid SharePoint reserved name conflicts
 */
interface ISharePointListItem {
  Id: number;
  Title: string;
  PD_UserPrincipalName: string;
  PD_Email?: string;
  PD_Department?: string;
  PD_JobTitle?: string;
  PD_OfficeLocation?: string;
  PD_BusinessPhones?: string;
  PD_MobilePhone?: string;
  PD_City?: string;
  PD_Country?: string;
  PD_CompanyName?: string;
  PD_PhotoUrl?: string;
  PD_GivenName?: string;
  PD_Surname?: string;
  PD_UserId?: string;
  PD_LastVerified?: string | Date;
  PD_AccessCount?: number;
}

/**
 * Sync metadata structure
 */
interface ISyncMetadata {
  lastFullSync: Date | null;
  totalUsers: number;
  lastSyncSuccess: boolean;
}

/**
 * Service for managing SharePoint list cache
 */
export class ListService {
  private sp: SPFI;
  private listTitle: string = Constants.LIST_TITLE;
  private isListReady = false;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Ensure the cache list exists, create if not
   */
  public async ensureList(): Promise<void> {
    if (this.isListReady) return;

    try {
      // Check if list exists
      try {
        await this.sp.web.lists.getByTitle(this.listTitle).select('Id')();
        this.isListReady = true;
        return;
      } catch (checkError) {
        // List doesn't exist, continue to create it
        console.log('List does not exist, will create it');
      }

      // Create the list
      try {
        const listAddResult = await this.sp.web.lists.add(this.listTitle, '', 100, false, {
          Hidden: false,
          OnQuickLaunch: false,
          AllowContentTypes: false
        });

        const list = listAddResult.list;

        // Add custom fields with PD_ prefix to avoid SharePoint reserved name conflicts
        await list.fields.addText('PD_UserPrincipalName', { MaxLength: 255, Required: true });
        await list.fields.addText('PD_Email', { MaxLength: 255 });
        await list.fields.addText('PD_Department', { MaxLength: 255 });
        await list.fields.addText('PD_JobTitle', { MaxLength: 255 });
        await list.fields.addText('PD_OfficeLocation', { MaxLength: 255 });
        await list.fields.addMultilineText('PD_BusinessPhones', { NumberOfLines: 2, RichText: false });
        await list.fields.addText('PD_MobilePhone', { MaxLength: 50 });
        await list.fields.addText('PD_City', { MaxLength: 100 });
        await list.fields.addText('PD_Country', { MaxLength: 100 });
        await list.fields.addText('PD_CompanyName', { MaxLength: 255 });
        await list.fields.addMultilineText('PD_PhotoUrl', { NumberOfLines: 2, RichText: false });
        await list.fields.addText('PD_GivenName', { MaxLength: 255 });
        await list.fields.addText('PD_Surname', { MaxLength: 255 });
        await list.fields.addText('PD_UserId', { MaxLength: 100 });
        await list.fields.addDateTime('PD_LastVerified', { DisplayFormat: 1 });
        await list.fields.addNumber('PD_AccessCount', { MinimumValue: 0 });

        // Create indexes for performance
        await list.fields.getByInternalNameOrTitle('PD_UserPrincipalName').update({ Indexed: true });
        await list.fields.getByInternalNameOrTitle('PD_Department').update({ Indexed: true });
        await list.fields.getByInternalNameOrTitle('PD_LastVerified').update({ Indexed: true });

        this.isListReady = true;
      } catch (createError: unknown) {
        // Check if error is "list already exists"
        const errorMessage = createError instanceof Error
          ? createError.message
          : String(createError);
        if (errorMessage.includes('already exists') || errorMessage.includes('-2130575342')) {
          console.log('List already exists, will use existing list');
          // List was created by another process/tab, just mark as ready
          this.isListReady = true;
          return;
        }
        // Other error, rethrow
        throw createError;
      }
    } catch (error) {
      throw new Error(ErrorHandler.getUserMessage(error, 'ListService.ensureList'));
    }
  }

  /**
   * Add or update user in cache list
   */
  public async upsertUser(user: IUserProfile): Promise<void> {
    try {
      await this.ensureList();

      const list = this.sp.web.lists.getByTitle(this.listTitle);

      // Check if user already exists
      const existingItems = await list.items
        .filter(`PD_UserPrincipalName eq '${user.userPrincipalName.replace(/'/g, "''")}'`)
        .top(1)
        .select('Id', 'PD_AccessCount')();

      const itemData = this.mapUserToListItem(user);

      if (existingItems.length > 0) {
        // Update existing item
        const existingItem = existingItems[0] as { Id: number; PD_AccessCount?: number };
        itemData.PD_AccessCount = (existingItem.PD_AccessCount || 0) + 1;
        await list.items.getById(existingItem.Id).update(itemData);
      } else {
        // Check if we're at capacity, remove least accessed item
        const items = await list.items.select('Id').top(1)();
        const itemCount = items.length > 0 ? await list.items.select('Id').top(5000)().then(i => i.length) : 0;
        if (itemCount >= Constants.LIST_MAX_ITEMS) {
          await this.evictLeastAccessedItem();
        }

        // Add new item
        itemData.PD_AccessCount = 1;
        await list.items.add(itemData);
      }
    } catch (error) {
      console.error('Error upserting user to list:', error);
      // Don't throw - cache updates should not break user experience
    }
  }

  /**
   * Get user from cache list
   */
  public async getUser(userPrincipalName: string): Promise<IUserProfile | null> {
    try {
      await this.ensureList();

      const items = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.filter(`PD_UserPrincipalName eq '${userPrincipalName.replace(/'/g, "''")}'`)
        .top(1)();

      if (items.length === 0) {
        return null;
      }

      const item = items[0];

      // Check if cache entry is expired
      const lastVerified = item.PD_LastVerified ? new Date(item.PD_LastVerified) : null;
      if (lastVerified && Date.now() - lastVerified.getTime() > Constants.LIST_CACHE_TTL) {
        return null;
      }

      return this.mapListItemToUser(item);
    } catch (error) {
      console.error('Error getting user from list:', error);
      return null;
    }
  }

  /**
   * Search users in cache list
   */
  public async searchUsers(searchText: string, top = 50): Promise<IUserProfile[]> {
    try {
      await this.ensureList();

      const escapedSearch = searchText.replace(/'/g, "''");
      const filter = `(substringof('${escapedSearch}', Title) or ` +
                    `substringof('${escapedSearch}', PD_Email) or ` +
                    `substringof('${escapedSearch}', PD_Department) or ` +
                    `substringof('${escapedSearch}', PD_JobTitle))`;

      const items = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.filter(filter)
        .top(top)
        .orderBy('PD_AccessCount', false)();

      return items.map(item => this.mapListItemToUser(item));
    } catch (error) {
      console.error('Error searching users in list:', error);
      return [];
    }
  }

  /**
   * Get all departments from cache
   */
  public async getDepartments(): Promise<string[]> {
    try {
      await this.ensureList();

      const items = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.select('PD_Department')
        .top(5000)();

      const departments = new Set<string>();
      items.forEach(item => {
        if (item.PD_Department) {
          departments.add(item.PD_Department);
        }
      });

      return Array.from(departments).sort();
    } catch (error) {
      console.error('Error getting departments from list:', error);
      return [];
    }
  }

  /**
   * Evict least recently accessed item from cache
   */
  private async evictLeastAccessedItem(): Promise<void> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.select('Id')
        .orderBy('PD_AccessCount', true)
        .top(1)();

      if (items.length > 0) {
        await this.sp.web.lists
          .getByTitle(this.listTitle)
          .items.getById(items[0].Id)
          .delete();
      }
    } catch (error) {
      console.error('Error evicting item from cache:', error);
    }
  }

  /**
   * Clear all cached items
   */
  public async clearCache(): Promise<void> {
    try {
      await this.ensureList();

      const items = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.select('Id')
        .top(5000)();

      // Delete items individually (batch API changed in v3)
      const deletePromises = items.map(item =>
        this.sp.web.lists.getByTitle(this.listTitle).items.getById(item.Id).delete()
      );
      await Promise.all(deletePromises);
    } catch (error) {
      throw new Error(ErrorHandler.getUserMessage(error, 'ListService.clearCache'));
    }
  }

  /**
   * Map user profile to SharePoint list item
   */
  private mapUserToListItem(user: IUserProfile): Partial<ISharePointListItem> {
    return {
      Title: user.displayName,
      PD_UserPrincipalName: user.userPrincipalName,
      PD_Email: user.mail || '',
      PD_Department: user.department || '',
      PD_JobTitle: user.jobTitle || '',
      PD_OfficeLocation: user.officeLocation || '',
      PD_BusinessPhones: user.businessPhones ? JSON.stringify(user.businessPhones) : '',
      PD_MobilePhone: user.mobilePhone || '',
      PD_City: user.city || '',
      PD_Country: user.country || '',
      PD_CompanyName: user.companyName || '',
      PD_PhotoUrl: user.photoUrl || '',
      PD_GivenName: user.givenName || '',
      PD_Surname: user.surname || '',
      PD_UserId: user.id,
      PD_LastVerified: user.lastVerified || new Date()
    };
  }

  /**
   * Get paginated users from list (for initial display)
   */
  public async getPaginatedUsers(pageSize = 30, pageNumber = 1, orderBy = 'Title'): Promise<IUserProfile[]> {
    try {
      await this.ensureList();

      const skipCount = (pageNumber - 1) * pageSize;

      const items = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.select('Id', 'PD_UserId', 'PD_UserPrincipalName', 'Title', 'PD_Email', 'PD_Department', 'PD_JobTitle',
          'PD_OfficeLocation', 'PD_BusinessPhones', 'PD_MobilePhone', 'PD_City', 'PD_Country', 'PD_CompanyName',
          'PD_PhotoUrl', 'PD_GivenName', 'PD_Surname', 'PD_LastVerified', 'PD_AccessCount')
        .orderBy(orderBy, true)
        .skip(skipCount)
        .top(pageSize)();

      return items.map(item => this.mapListItemToUser(item));
    } catch (error) {
      console.error('Error getting paginated users:', error);
      return [];
    }
  }

  /**
   * Get total user count from list
   */
  public async getTotalUserCount(): Promise<number> {
    try {
      await this.ensureList();

      // For large lists, we need to estimate
      // This is a limitation of SharePoint - getting exact count > 5000 is expensive
      const result = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.select('Id')
        .top(5000)();

      return result.length;
    } catch (error) {
      console.error('Error getting user count:', error);
      return 0;
    }
  }

  /**
   * Add or update user (optimized for bulk operations)
   */
  public async addOrUpdateUser(user: IUserProfile): Promise<void> {
    try {
      await this.ensureList();

      const list = this.sp.web.lists.getByTitle(this.listTitle);
      const itemData = this.mapUserToListItem(user);

      // Check if user exists
      const existingItems = await list.items
        .filter(`PD_UserPrincipalName eq '${user.userPrincipalName.replace(/'/g, "''")}'`)
        .top(1)
        .select('Id')();

      if (existingItems.length > 0) {
        // Update existing
        await list.items.getById(existingItems[0].Id).update(itemData);
      } else {
        // Add new
        await list.items.add(itemData);
      }
    } catch (error) {
      console.error(`Error adding/updating user ${user.userPrincipalName}:`, error);
      throw error;
    }
  }

  /**
   * Get sync metadata
   */
  public async getSyncMetadata(): Promise<ISyncMetadata | null> {
    try {
      await this.ensureList();

      // Try to get metadata from a special "metadata" item (Title = "_SyncMetadata")
      const items = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.filter("Title eq '_SyncMetadata'")
        .top(1)();

      if (items.length === 0) {
        return null;
      }

      return {
        lastFullSync: items[0].PD_LastVerified ? new Date(items[0].PD_LastVerified) : null,
        totalUsers: items[0].PD_AccessCount || 0,
        lastSyncSuccess: items[0].PD_Department === 'Success'
      };
    } catch (error) {
      console.error('Error getting sync metadata:', error);
      return null;
    }
  }

  /**
   * Update sync metadata
   */
  public async updateSyncMetadata(metadata: { lastFullSync: Date; totalUsers: number; lastSyncSuccess: boolean }): Promise<void> {
    try {
      await this.ensureList();

      const list = this.sp.web.lists.getByTitle(this.listTitle);

      // Check if metadata item exists
      const existingItems = await list.items
        .filter("Title eq '_SyncMetadata'")
        .top(1)
        .select('Id')();

      const metadataItem = {
        Title: '_SyncMetadata',
        PD_UserPrincipalName: 'system',
        PD_LastVerified: metadata.lastFullSync,
        PD_AccessCount: metadata.totalUsers,
        PD_Department: metadata.lastSyncSuccess ? 'Success' : 'Failed',
        PD_Email: 'Sync Metadata - Do Not Delete'
      };

      if (existingItems.length > 0) {
        // Update existing metadata
        await list.items.getById(existingItems[0].Id).update(metadataItem);
      } else {
        // Create metadata item
        await list.items.add(metadataItem);
      }
    } catch (error) {
      console.error('Error updating sync metadata:', error);
      // Don't throw - metadata update failure shouldn't break the sync
    }
  }

  /**
   * Get advanced filtered users
   */
  public async getFilteredUsers(filters: {
    department?: string;
    officeLocation?: string;
    city?: string;
    country?: string;
    jobTitle?: string;
  }, pageSize = 50): Promise<IUserProfile[]> {
    try {
      await this.ensureList();

      let filterQuery = "Title ne '_SyncMetadata'"; // Exclude metadata item

      if (filters.department) {
        filterQuery += ` and PD_Department eq '${filters.department.replace(/'/g, "''")}'`;
      }
      if (filters.officeLocation) {
        filterQuery += ` and PD_OfficeLocation eq '${filters.officeLocation.replace(/'/g, "''")}'`;
      }
      if (filters.city) {
        filterQuery += ` and PD_City eq '${filters.city.replace(/'/g, "''")}'`;
      }
      if (filters.country) {
        filterQuery += ` and PD_Country eq '${filters.country.replace(/'/g, "''")}'`;
      }
      if (filters.jobTitle) {
        filterQuery += ` and substringof('${filters.jobTitle.replace(/'/g, "''")}', PD_JobTitle)`;
      }

      const items = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.filter(filterQuery)
        .top(pageSize)
        .orderBy('Title', true)();

      return items.map(item => this.mapListItemToUser(item));
    } catch (error) {
      console.error('Error getting filtered users:', error);
      return [];
    }
  }

  /**
   * Map SharePoint list item to user profile
   */
  private mapListItemToUser(item: ISharePointListItem): IUserProfile {
    return {
      id: item.PD_UserId,
      userPrincipalName: item.PD_UserPrincipalName,
      displayName: item.Title,
      givenName: item.PD_GivenName,
      surname: item.PD_Surname,
      mail: item.PD_Email,
      jobTitle: item.PD_JobTitle,
      department: item.PD_Department,
      officeLocation: item.PD_OfficeLocation,
      businessPhones: item.PD_BusinessPhones ? JSON.parse(item.PD_BusinessPhones) : [],
      mobilePhone: item.PD_MobilePhone,
      city: item.PD_City,
      country: item.PD_Country,
      companyName: item.PD_CompanyName,
      photoUrl: item.PD_PhotoUrl,
      lastVerified: item.PD_LastVerified ? new Date(item.PD_LastVerified) : undefined,
      accessCount: item.PD_AccessCount || 0
    };
  }
}
