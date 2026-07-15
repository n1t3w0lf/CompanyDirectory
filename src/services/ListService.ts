import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import { IList } from '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/batching';
import { IUserProfile } from '../models/IUserProfile';
import { Constants } from '../models/Constants';
import { ErrorHandler } from '../utils/ErrorHandler';

/**
 * Cache-list schema: the single source of truth for the custom columns.
 * Used both to create the list and to repair an existing list that is missing
 * columns (schema drift / interrupted creation).
 */
type FieldKind = 'text' | 'multiline' | 'datetime' | 'number';
interface IFieldDef {
  name: string;
  kind: FieldKind;
  maxLength?: number;
  required?: boolean;
}

const REQUIRED_FIELDS: IFieldDef[] = [
  { name: 'PD_UserPrincipalName', kind: 'text', maxLength: 255, required: true },
  { name: 'PD_Email', kind: 'text', maxLength: 255 },
  { name: 'PD_Department', kind: 'text', maxLength: 255 },
  { name: 'PD_JobTitle', kind: 'text', maxLength: 255 },
  { name: 'PD_OfficeLocation', kind: 'text', maxLength: 255 },
  { name: 'PD_BusinessPhones', kind: 'multiline' },
  { name: 'PD_MobilePhone', kind: 'text', maxLength: 50 },
  { name: 'PD_City', kind: 'text', maxLength: 100 },
  { name: 'PD_Country', kind: 'text', maxLength: 100 },
  { name: 'PD_CompanyName', kind: 'text', maxLength: 255 },
  { name: 'PD_PhotoUrl', kind: 'multiline' },
  { name: 'PD_GivenName', kind: 'text', maxLength: 255 },
  { name: 'PD_Surname', kind: 'text', maxLength: 255 },
  { name: 'PD_UserId', kind: 'text', maxLength: 100 },
  { name: 'PD_LastVerified', kind: 'datetime' },
  { name: 'PD_AccessCount', kind: 'number' }
];

// Columns to index. Covers exact lookups, `eq` filters and `orderBy` targets —
// NOT the `substringof` (contains) search columns, which SharePoint cannot serve
// from an index. NOTE: SharePoint can only build an index while the list is
// under the 5,000-item threshold, so index before a large sync populates it;
// ensureFields applies these best-effort (failures on an already-large list are
// logged, not fatal). Max 20 indexed columns per list.
const INDEXED_FIELDS = [
  'PD_UserPrincipalName', // exact lookup (dedupe / upsert)
  'PD_Department',        // eq filter + search
  'PD_LastVerified',      // retained
  'Title',                // orderBy (initial list + filtered results)
  'PD_OfficeLocation',    // eq filter
  'PD_City',              // eq filter
  'PD_Country',           // eq filter
  'PD_AccessCount'        // orderBy (search ranking)
];

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
        // List exists — ensure its schema is complete (repairs missing columns
        // from an interrupted creation or an older build).
        await this.ensureFields(this.sp.web.lists.getByTitle(this.listTitle));
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

        await this.ensureFields(listAddResult.list);
        this.isListReady = true;
      } catch (createError: unknown) {
        // Check if error is "list already exists" (created by another tab/process)
        const errorMessage = createError instanceof Error
          ? createError.message
          : String(createError);
        if (errorMessage.includes('already exists') || errorMessage.includes('-2130575342')) {
          console.log('List already exists, ensuring schema on existing list');
          await this.ensureFields(this.sp.web.lists.getByTitle(this.listTitle));
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
   * Idempotently ensure every required column exists on the list, adding any
   * that are missing (schema self-heal). Reads existing internal names once and
   * only adds the gaps, so it is cheap when the schema is already complete.
   * Also re-applies indexing (idempotent) on the key columns.
   */
  private async ensureFields(list: IList): Promise<void> {
    const existing = new Set<string>();
    try {
      const fields: Array<{ InternalName: string }> = await list.fields.select('InternalName').top(500)();
      fields.forEach(f => existing.add(f.InternalName));
    } catch (error) {
      console.warn('ensureFields: could not read existing columns; will attempt all adds:', error);
    }

    for (const def of REQUIRED_FIELDS) {
      if (existing.has(def.name)) {
        continue;
      }
      try {
        switch (def.kind) {
          case 'text':
            await list.fields.addText(def.name, { MaxLength: def.maxLength || 255, Required: def.required === true });
            break;
          case 'multiline':
            await list.fields.addMultilineText(def.name, { NumberOfLines: 2, RichText: false });
            break;
          case 'datetime':
            await list.fields.addDateTime(def.name, { DisplayFormat: 1 });
            break;
          case 'number':
            await list.fields.addNumber(def.name, { MinimumValue: 0 });
            break;
        }
      } catch (error) {
        // A concurrent creator may have added it between our read and write.
        const msg = error instanceof Error ? error.message : String(error);
        if (!/exists|duplicate/i.test(msg)) {
          console.warn(`ensureFields: failed to add column ${def.name}:`, error);
        }
      }
    }

    // Best-effort, idempotent indexing on the key query columns.
    for (const name of INDEXED_FIELDS) {
      try {
        await list.fields.getByInternalNameOrTitle(name).update({ Indexed: true });
      } catch (error) {
        console.warn(`ensureFields: could not index ${name}:`, error);
      }
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
      // Contains-match across the searchable columns; exclude the metadata row.
      const filter = `Title ne '_SyncMetadata' and ` +
                    `(substringof('${escapedSearch}', Title) or ` +
                    `substringof('${escapedSearch}', PD_Email) or ` +
                    `substringof('${escapedSearch}', PD_Department) or ` +
                    `substringof('${escapedSearch}', PD_JobTitle) or ` +
                    `substringof('${escapedSearch}', PD_GivenName) or ` +
                    `substringof('${escapedSearch}', PD_Surname))`;

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
   * Get total user count from list. Uses the list's ItemCount property, which is
   * accurate beyond the 5,000 view threshold and costs a single cheap call.
   * Subtracts the `_SyncMetadata` bookkeeping row.
   */
  public async getTotalUserCount(): Promise<number> {
    try {
      await this.ensureList();

      const listInfo = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .select('ItemCount')();

      const count = (listInfo as { ItemCount?: number }).ItemCount || 0;
      // Exclude the single _SyncMetadata bookkeeping row when present.
      return count > 0 ? count - 1 : 0;
    } catch (error) {
      console.error('Error getting user count:', error);
      return 0;
    }
  }

  /**
   * Bulk-read every existing row's SharePoint item Id keyed by lowercased UPN,
   * using keyset paging on the indexed Id column so it scales past the 5,000
   * list-view threshold. Used by the sync to route add-vs-update in memory and
   * eliminate the per-user existence query.
   */
  public async getExistingUserMap(): Promise<Map<string, number>> {
    await this.ensureList();

    const map = new Map<string, number>();
    const pageSize = Constants.LIST_PAGE_SIZE;
    let lastId = 0;
    let page: Array<{ Id: number; PD_UserPrincipalName?: string }>;

    do {
      page = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.select('Id', 'PD_UserPrincipalName')
        .filter(`Id gt ${lastId}`)
        .orderBy('Id', true)
        .top(pageSize)();

      for (const item of page) {
        if (item.PD_UserPrincipalName && item.PD_UserPrincipalName.toLowerCase() !== 'system') {
          map.set(item.PD_UserPrincipalName.toLowerCase(), item.Id);
        }
        if (item.Id > lastId) {
          lastId = item.Id;
        }
      }
    } while (page.length === pageSize);

    return map;
  }

  /**
   * Bulk add/update users via SharePoint $batch, routing add-vs-update from a
   * preloaded UPN->Id map (so no per-user existence query, and no duplicates on
   * re-run). Input is de-duped by UPN. Throttled/failed items are retried with
   * Retry-After backoff rather than silently dropped.
   */
  public async bulkUpsertUsers(
    users: IUserProfile[],
    existingMap: Map<string, number>,
    onProgress?: (processed: number) => void,
    isCancelled?: () => boolean
  ): Promise<{ added: number; updated: number; failed: number }> {
    await this.ensureList();

    // De-dupe by lowercased UPN (last wins) to guard the no-duplicate invariant.
    const byUpn = new Map<string, IUserProfile>();
    for (const user of users) {
      if (user.userPrincipalName) {
        byUpn.set(user.userPrincipalName.toLowerCase(), user);
      }
    }
    const unique = Array.from(byUpn.values());

    let added = 0;
    let updated = 0;
    let failed = 0;
    let processed = 0;

    for (let i = 0; i < unique.length; i += Constants.LIST_BATCH_SIZE) {
      if (isCancelled && isCancelled()) {
        break;
      }

      const chunk = unique.slice(i, i + Constants.LIST_BATCH_SIZE);
      const result = await this.writeChunkWithRetry(chunk, existingMap, isCancelled);
      added += result.added;
      updated += result.updated;
      failed += result.failed;

      processed += chunk.length;
      if (onProgress) {
        onProgress(processed);
      }
    }

    return { added, updated, failed };
  }

  /**
   * Write one chunk as a single $batch, retrying throttled items (honouring
   * Retry-After) up to Constants.SYNC_MAX_RETRIES before giving up on them.
   */
  private async writeChunkWithRetry(
    chunk: IUserProfile[],
    existingMap: Map<string, number>,
    isCancelled?: () => boolean
  ): Promise<{ added: number; updated: number; failed: number }> {
    let pending = chunk;
    let added = 0;
    let updated = 0;
    let failed = 0;

    for (let attempt = 0; attempt <= Constants.SYNC_MAX_RETRIES && pending.length > 0; attempt++) {
      if (isCancelled && isCancelled()) {
        break;
      }

      const [batchedSp, execute] = this.sp.batched();
      const list = batchedSp.web.lists.getByTitle(this.listTitle);

      const ops = pending.map(user => {
        const itemData = this.mapUserToListItem(user);
        const existingId = existingMap.get(user.userPrincipalName.toLowerCase());
        const op = existingId
          ? list.items.getById(existingId).update(itemData)
          : list.items.add(itemData);
        return op.then(
          () => ({ ok: true, existing: !!existingId, user, error: undefined as unknown }),
          (error: unknown) => ({ ok: false, existing: !!existingId, user, error })
        );
      });

      try {
        await execute();
      } catch (error) {
        // The $batch POST itself failed. PnP only settles the per-op promises on
        // the POST success path, so `ops` will NEVER settle here — do NOT await
        // them. Retry the whole chunk on throttling, otherwise abandon it.
        if (this.isThrottleError(error) && attempt < Constants.SYNC_MAX_RETRIES) {
          await this.delay(this.getRetryAfterMs(error, attempt));
          continue; // pending unchanged -> retry the whole chunk
        }
        console.error('Bulk upsert batch POST failed, abandoning chunk:', error);
        break; // remaining `pending` is counted as failed below
      }

      // Batch POST succeeded: inspect per-op results.
      const results = await Promise.all(ops);
      const nextPending: IUserProfile[] = [];
      let retryAfterMs = 0;
      for (const r of results) {
        if (r.ok) {
          if (r.existing) { updated++; } else { added++; }
        } else if (this.isThrottleError(r.error)) {
          nextPending.push(r.user);
          retryAfterMs = Math.max(retryAfterMs, this.getRetryAfterMs(r.error, attempt));
        } else {
          failed++; // permanent per-item failure: count it, do not retry
          console.error(`Bulk upsert failed for ${r.user.userPrincipalName}:`, r.error);
        }
      }

      pending = nextPending;
      if (pending.length > 0 && attempt < Constants.SYNC_MAX_RETRIES) {
        await this.delay(retryAfterMs || this.backoffMs(attempt));
      }
    }

    // Anything still pending (throttled through all retries) plus permanent failures.
    return { added, updated, failed: failed + pending.length };
  }

  /**
   * True if the error looks like SharePoint throttling (429/503).
   */
  private isThrottleError(error: unknown): boolean {
    if (!error) {
      return false;
    }
    const status = (error as { status?: number; statusCode?: number }).status
      || (error as { statusCode?: number }).statusCode;
    if (status === 429 || status === 503) {
      return true;
    }
    const text = String((error as { message?: string }).message || error);
    return text.indexOf('429') >= 0 || text.indexOf('503') >= 0 || text.indexOf('Too Many Requests') >= 0;
  }

  /**
   * Milliseconds to wait before retrying, honouring a Retry-After header if the
   * error carries a response, else an exponential backoff with jitter.
   */
  private getRetryAfterMs(error: unknown, attempt: number): number {
    try {
      const response = (error as { response?: { headers?: { get?: (name: string) => string | null } } }).response;
      const header = response && response.headers && response.headers.get
        ? response.headers.get('Retry-After')
        : null;
      if (header) {
        const seconds = parseInt(header, 10);
        if (!isNaN(seconds) && seconds > 0) {
          return seconds * 1000;
        }
      }
    } catch {
      // fall through to backoff
    }
    return this.backoffMs(attempt);
  }

  /**
   * Exponential backoff with jitter (base 1s, capped ~30s).
   */
  private backoffMs(attempt: number): number {
    const base = Math.min(30000, 1000 * Math.pow(2, attempt));
    return base + Math.floor(Math.random() * base * 0.2);
  }

  /**
   * Delay helper.
   */
  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
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
