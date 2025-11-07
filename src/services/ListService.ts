import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import { IUserProfile } from '../models/IUserProfile';
import { Constants } from '../models/Constants';
import { ErrorHandler } from '../utils/ErrorHandler';

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
      let list;
      let listExists = false;

      // Check if list exists
      try {
        list = this.sp.web.lists.getByTitle(this.listTitle);
        await list.select('Id')();
        listExists = true;
      } catch {
        // List doesn't exist, create it
        listExists = false;
      }

      if (!listExists) {
        // Create the list
        const listAddResult = await this.sp.web.lists.add(this.listTitle, '', 100, false, {
          Hidden: false,
          OnQuickLaunch: false,
          AllowContentTypes: false
        });
        list = listAddResult.list;
      }

      // Ensure all required fields exist
      await this.ensureFields(list);

      this.isListReady = true;
    } catch (error) {
      throw new Error(ErrorHandler.getUserMessage(error, 'ListService.ensureList'));
    }
  }

  /**
   * Ensure all required fields exist in the list
   */
  private async ensureFields(list: any): Promise<void> {
    try {
      // Get existing fields
      const existingFields = await list.fields.select('InternalName')();
      const existingFieldNames = new Set(existingFields.map((f: any) => f.InternalName));

      // Define all required fields
      const requiredFields = [
        { name: 'UserPrincipalName', type: 'text', options: { MaxLength: 255, Required: true } },
        { name: 'Email', type: 'text', options: { MaxLength: 255 } },
        { name: 'Department', type: 'text', options: { MaxLength: 255 } },
        { name: 'JobTitle', type: 'text', options: { MaxLength: 255 } },
        { name: 'OfficeLocation', type: 'text', options: { MaxLength: 255 } },
        { name: 'BusinessPhones', type: 'text', options: { MaxLength: 500 } },
        { name: 'MobilePhone', type: 'text', options: { MaxLength: 50 } },
        { name: 'City', type: 'text', options: { MaxLength: 100 } },
        { name: 'Country', type: 'text', options: { MaxLength: 100 } },
        { name: 'CompanyName', type: 'text', options: { MaxLength: 255 } },
        { name: 'PhotoUrl', type: 'text', options: { MaxLength: 1000 } },
        { name: 'GivenName', type: 'text', options: { MaxLength: 255 } },
        { name: 'Surname', type: 'text', options: { MaxLength: 255 } },
        { name: 'UserId', type: 'text', options: { MaxLength: 100 } },
        { name: 'LastVerified', type: 'datetime', options: { DisplayFormat: 1 } },
        { name: 'AccessCount', type: 'number', options: { MinimumValue: 0 } }
      ];

      // Add missing fields
      for (const field of requiredFields) {
        if (!existingFieldNames.has(field.name)) {
          console.log(`Adding missing field: ${field.name}`);
          try {
            if (field.type === 'text') {
              await list.fields.addText(field.name, field.options);
            } else if (field.type === 'datetime') {
              await list.fields.addDateTime(field.name, field.options);
            } else if (field.type === 'number') {
              await list.fields.addNumber(field.name, field.options);
            }
          } catch (fieldError) {
            console.warn(`Error adding field ${field.name}:`, fieldError);
            // Continue with other fields even if one fails
          }
        }
      }

      // Create indexes for performance (only if fields don't have indexes)
      const indexFields = ['UserPrincipalName', 'Department', 'LastVerified'];
      for (const fieldName of indexFields) {
        try {
          if (existingFieldNames.has(fieldName)) {
            await list.fields.getByInternalNameOrTitle(fieldName).update({ Indexed: true });
          }
        } catch (indexError) {
          console.warn(`Could not create index on ${fieldName}:`, indexError);
          // Index creation is optional, continue anyway
        }
      }
    } catch (error) {
      console.error('Error ensuring fields:', error);
      // Don't throw - we'll try to work with whatever fields exist
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
        .filter(`UserPrincipalName eq '${user.userPrincipalName.replace(/'/g, "''")}'`)
        .top(1)
        .select('Id', 'AccessCount')();

      const itemData = this.mapUserToListItem(user);

      if (existingItems.length > 0) {
        // Update existing item
        const existingItem = existingItems[0];
        itemData.AccessCount = (existingItem.AccessCount || 0) + 1;
        await list.items.getById(existingItem.Id).update(itemData);
      } else {
        // Check if we're at capacity, remove least accessed item
        const items = await list.items.select('Id').top(1)();
        const itemCount = items.length > 0 ? await list.items.select('Id').top(5000)().then(i => i.length) : 0;
        if (itemCount >= Constants.LIST_MAX_ITEMS) {
          await this.evictLeastAccessedItem();
        }

        // Add new item
        itemData.AccessCount = 1;
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
        .items.filter(`UserPrincipalName eq '${userPrincipalName.replace(/'/g, "''")}'`)
        .top(1)();

      if (items.length === 0) {
        return null;
      }

      const item = items[0];

      // Check if cache entry is expired
      const lastVerified = item.LastVerified ? new Date(item.LastVerified) : null;
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
                    `substringof('${escapedSearch}', Email) or ` +
                    `substringof('${escapedSearch}', Department) or ` +
                    `substringof('${escapedSearch}', JobTitle))`;

      const items = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.filter(filter)
        .top(top)
        .orderBy('AccessCount', false)();

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
        .items.select('Department')
        .top(5000)();

      const departments = new Set<string>();
      items.forEach(item => {
        if (item.Department) {
          departments.add(item.Department);
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
        .orderBy('AccessCount', true)
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
  private mapUserToListItem(user: IUserProfile): any {
    return {
      Title: user.displayName,
      UserPrincipalName: user.userPrincipalName,
      Email: user.mail || '',
      Department: user.department || '',
      JobTitle: user.jobTitle || '',
      OfficeLocation: user.officeLocation || '',
      BusinessPhones: user.businessPhones ? JSON.stringify(user.businessPhones) : '',
      MobilePhone: user.mobilePhone || '',
      City: user.city || '',
      Country: user.country || '',
      CompanyName: user.companyName || '',
      PhotoUrl: user.photoUrl || '',
      GivenName: user.givenName || '',
      Surname: user.surname || '',
      UserId: user.id,
      LastVerified: user.lastVerified || new Date()
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
        .items.select('Id', 'UserId', 'UserPrincipalName', 'Title', 'Email', 'Department', 'JobTitle',
          'OfficeLocation', 'BusinessPhones', 'MobilePhone', 'City', 'Country', 'CompanyName',
          'PhotoUrl', 'GivenName', 'Surname', 'LastVerified', 'AccessCount')
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

      // Use a simple count query
      const items = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.select('Id')
        .top(1)();

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
        .filter(`UserPrincipalName eq '${user.userPrincipalName.replace(/'/g, "''")}'`)
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
  public async getSyncMetadata(): Promise<any> {
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
        lastFullSync: items[0].LastVerified ? new Date(items[0].LastVerified) : null,
        totalUsers: items[0].AccessCount || 0,
        lastSyncSuccess: items[0].Department === 'Success'
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
        UserPrincipalName: 'system',
        LastVerified: metadata.lastFullSync,
        AccessCount: metadata.totalUsers,
        Department: metadata.lastSyncSuccess ? 'Success' : 'Failed',
        Email: 'Sync Metadata - Do Not Delete'
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
        filterQuery += ` and Department eq '${filters.department.replace(/'/g, "''")}'`;
      }
      if (filters.officeLocation) {
        filterQuery += ` and OfficeLocation eq '${filters.officeLocation.replace(/'/g, "''")}'`;
      }
      if (filters.city) {
        filterQuery += ` and City eq '${filters.city.replace(/'/g, "''")}'`;
      }
      if (filters.country) {
        filterQuery += ` and Country eq '${filters.country.replace(/'/g, "''")}'`;
      }
      if (filters.jobTitle) {
        filterQuery += ` and substringof('${filters.jobTitle.replace(/'/g, "''")}', JobTitle)`;
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
  private mapListItemToUser(item: any): IUserProfile {
    return {
      id: item.UserId,
      userPrincipalName: item.UserPrincipalName,
      displayName: item.Title,
      givenName: item.GivenName,
      surname: item.Surname,
      mail: item.Email,
      jobTitle: item.JobTitle,
      department: item.Department,
      officeLocation: item.OfficeLocation,
      businessPhones: item.BusinessPhones ? JSON.parse(item.BusinessPhones) : [],
      mobilePhone: item.MobilePhone,
      city: item.City,
      country: item.Country,
      companyName: item.CompanyName,
      photoUrl: item.PhotoUrl,
      lastVerified: item.LastVerified ? new Date(item.LastVerified) : undefined,
      accessCount: item.AccessCount || 0
    };
  }
}
