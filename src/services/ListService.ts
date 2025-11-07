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
  private isListReady: boolean = false;

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
      } catch {
        // List doesn't exist, create it
      }

      // Create the list
      const listAddResult = await this.sp.web.lists.add(this.listTitle, '', 100, false, {
        Hidden: false,
        OnQuickLaunch: false,
        AllowContentTypes: false
      });

      const list = listAddResult.list;

      // Add custom fields
      await list.fields.addText('UserPrincipalName', { MaxLength: 255, Required: true });
      await list.fields.addText('Email', { MaxLength: 255 });
      await list.fields.addText('Department', { MaxLength: 255 });
      await list.fields.addText('JobTitle', { MaxLength: 255 });
      await list.fields.addText('OfficeLocation', { MaxLength: 255 });
      await list.fields.addText('BusinessPhones', { MaxLength: 500 });
      await list.fields.addText('MobilePhone', { MaxLength: 50 });
      await list.fields.addText('City', { MaxLength: 100 });
      await list.fields.addText('Country', { MaxLength: 100 });
      await list.fields.addText('CompanyName', { MaxLength: 255 });
      await list.fields.addText('PhotoUrl', { MaxLength: 1000 });
      await list.fields.addText('GivenName', { MaxLength: 255 });
      await list.fields.addText('Surname', { MaxLength: 255 });
      await list.fields.addText('UserId', { MaxLength: 100 });
      await list.fields.addDateTime('LastVerified', { DisplayFormat: 1 });
      await list.fields.addNumber('AccessCount', { MinimumValue: 0 });

      // Create indexes for performance
      await list.fields.getByInternalNameOrTitle('UserPrincipalName').update({ Indexed: true });
      await list.fields.getByInternalNameOrTitle('Department').update({ Indexed: true });
      await list.fields.getByInternalNameOrTitle('LastVerified').update({ Indexed: true });

      this.isListReady = true;
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
        const itemCount = await list.itemCount();
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
  public async searchUsers(searchText: string, top: number = 50): Promise<IUserProfile[]> {
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

      // Delete in batches
      const batch = this.sp.web.createBatch();
      items.forEach(item => {
        this.sp.web.lists.getByTitle(this.listTitle).items.getById(item.Id).inBatch(batch).delete();
      });
      await batch.execute();
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
