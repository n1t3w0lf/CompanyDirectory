import { GraphService } from './GraphService';
import { ListService } from './ListService';
import { IUserProfile } from '../models/IUserProfile';
import { ErrorHandler } from '../utils/ErrorHandler';

/**
 * Sync status for tracking bulk operations
 */
export interface ISyncStatus {
  isRunning: boolean;
  totalUsers: number;
  processedUsers: number;
  currentBatch: number;
  totalBatches: number;
  lastSyncDate?: Date;
  errors: string[];
  progress: number; // 0-100
}

/**
 * Service for bulk synchronization of users from Entra ID to SharePoint list
 */
export class SyncService {
  private graphService: GraphService;
  private listService: ListService;
  private syncStatus: ISyncStatus;
  private readonly BATCH_SIZE = 20; // Process 20 users at a time (conservative for throttling)
  private readonly GRAPH_PAGE_SIZE = 999; // Max Graph API page size
  private readonly BASE_DELAY = 2500; // 2.5 seconds between batches
  private readonly ERROR_DELAY = 8000; // 8 seconds after errors
  private cancelRequested = false;
  private consecutiveThrottles = 0;

  constructor(graphService: GraphService, listService: ListService) {
    this.graphService = graphService;
    this.listService = listService;
    this.syncStatus = this.getDefaultSyncStatus();
  }

  /**
   * Get current sync status
   */
  public getSyncStatus(): ISyncStatus {
    return { ...this.syncStatus };
  }

  /**
   * Check if initial sync is needed
   */
  public async isInitialSyncNeeded(): Promise<boolean> {
    try {
      const syncMetadata = await this.listService.getSyncMetadata();
      return !syncMetadata || !syncMetadata.lastFullSync;
    } catch (error) {
      console.error('Error checking sync status:', error);
      return true; // Assume sync needed if we can't check
    }
  }

  /**
   * Perform initial bulk sync of all users from Entra ID
   */
  public async performInitialSync(progressCallback?: (status: ISyncStatus) => void): Promise<void> {
    if (this.syncStatus.isRunning) {
      throw new Error('Sync is already running');
    }

    try {
      this.syncStatus.isRunning = true;
      this.syncStatus.errors = [];
      this.syncStatus.processedUsers = 0;
      this.syncStatus.currentBatch = 0;
      this.cancelRequested = false;
      this.consecutiveThrottles = 0;

      console.log('Starting initial sync of all users from Entra ID...');

      // Ensure list is created
      await this.listService.ensureList();

      // Fetch all users from Graph API with pagination
      const allUsers = await this.fetchAllUsersFromGraph(progressCallback);

      if (this.cancelRequested) {
        console.log('Sync cancelled during user fetch');
        return;
      }

      this.syncStatus.totalUsers = allUsers.length;
      this.syncStatus.totalBatches = Math.ceil(allUsers.length / this.BATCH_SIZE);
      this.syncStatus.processedUsers = 0; // Reset for batch processing phase

      console.log(`Fetched ${allUsers.length} users from Entra ID. Starting list population...`);

      // Reset throttle counter
      this.consecutiveThrottles = 0;

      // Process users in batches to avoid overwhelming SharePoint
      for (let i = 0; i < allUsers.length; i += this.BATCH_SIZE) {
        // Check for cancellation
        if (this.cancelRequested) {
          console.log('Sync cancelled. Saving progress...');
          break;
        }

        const batch = allUsers.slice(i, i + this.BATCH_SIZE);
        this.syncStatus.currentBatch = Math.floor(i / this.BATCH_SIZE) + 1;

        try {
          await this.processBatch(batch);
          this.syncStatus.processedUsers += batch.length;
          this.syncStatus.progress = Math.round((this.syncStatus.processedUsers / this.syncStatus.totalUsers) * 100);

          // Reset throttle counter on success
          this.consecutiveThrottles = 0;

          if (progressCallback) {
            progressCallback(this.getSyncStatus());
          }

          console.log(`Processed batch ${this.syncStatus.currentBatch}/${this.syncStatus.totalBatches} (${this.syncStatus.processedUsers}/${this.syncStatus.totalUsers} users)`);

          // Base delay between batches to avoid SharePoint throttling
          await this.delay(this.BASE_DELAY);
        } catch (error) {
          const errorMsg = `Error processing batch ${this.syncStatus.currentBatch}: ${ErrorHandler.getUserMessage(error)}`;
          this.syncStatus.errors.push(errorMsg);
          console.error(errorMsg);

          // Check if it's a throttling error
          const errorString = String(error);
          if (errorString.includes('429') || errorString.includes('Too Many Requests') ||
              errorString.includes('406') || errorString.includes('Throttle')) {
            this.consecutiveThrottles++;
            // Progressive backoff for throttling: multiply delay by number of consecutive throttles
            const throttleDelay = this.ERROR_DELAY * Math.min(this.consecutiveThrottles, 3);
            console.warn(`Throttled ${this.consecutiveThrottles} times. Waiting ${throttleDelay}ms before retry...`);
            await this.delay(throttleDelay);
          } else {
            // Non-throttling error, use standard error delay
            await this.delay(this.ERROR_DELAY);
          }
          // Continue with next batch even if one fails
        }
      }

      // Update sync metadata with progress
      await this.listService.updateSyncMetadata({
        lastFullSync: new Date(),
        totalUsers: this.syncStatus.processedUsers, // Use processed count, not total
        lastSyncSuccess: this.syncStatus.errors.length === 0 && !this.cancelRequested
      });

      this.syncStatus.lastSyncDate = new Date();
      this.syncStatus.progress = 100;

      if (progressCallback) {
        progressCallback(this.getSyncStatus());
      }

      console.log(`Initial sync completed. ${this.syncStatus.processedUsers} users synced with ${this.syncStatus.errors.length} errors.`);

    } catch (error) {
      const errorMsg = ErrorHandler.getUserMessage(error, 'SyncService.performInitialSync');
      this.syncStatus.errors.push(errorMsg);

      // Try to save whatever progress we made
      try {
        await this.listService.updateSyncMetadata({
          lastFullSync: new Date(),
          totalUsers: this.syncStatus.processedUsers,
          lastSyncSuccess: false
        });
      } catch (metadataError) {
        console.error('Failed to save sync metadata:', metadataError);
      }

      throw new Error(errorMsg);
    } finally {
      this.syncStatus.isRunning = false;
      this.cancelRequested = false;
    }
  }

  /**
   * Perform incremental sync (update existing users)
   */
  public async performIncrementalSync(userIds: string[]): Promise<void> {
    if (this.syncStatus.isRunning) {
      throw new Error('Sync is already running');
    }

    try {
      this.syncStatus.isRunning = true;
      this.syncStatus.errors = [];
      this.syncStatus.totalUsers = userIds.length;
      this.syncStatus.processedUsers = 0;

      console.log(`Starting incremental sync for ${userIds.length} users...`);

      // Fetch users in batches
      const users = await this.graphService.getUsersBatch(userIds);

      // Update in SharePoint list
      for (const user of users) {
        try {
          await this.listService.upsertUser(user);
          this.syncStatus.processedUsers++;
        } catch (error) {
          this.syncStatus.errors.push(`Error updating user ${user.userPrincipalName}: ${error}`);
        }
      }

      console.log(`Incremental sync completed. ${this.syncStatus.processedUsers} users updated.`);

    } catch (error) {
      const errorMsg = ErrorHandler.getUserMessage(error, 'SyncService.performIncrementalSync');
      this.syncStatus.errors.push(errorMsg);
      throw new Error(errorMsg);
    } finally {
      this.syncStatus.isRunning = false;
    }
  }

  /**
   * Fetch all users from Graph API with pagination
   */
  private async fetchAllUsersFromGraph(progressCallback?: (status: ISyncStatus) => void): Promise<IUserProfile[]> {
    const allUsers: IUserProfile[] = [];
    let nextLink: string | undefined;
    let pageCount = 0;
    let consecutiveErrors = 0;
    const MAX_CONSECUTIVE_ERRORS = 3;

    do {
      // Check for cancellation
      if (this.cancelRequested) {
        console.log('Fetch cancelled by user');
        break;
      }

      try {
        const result = await this.graphService.searchUsers('', this.GRAPH_PAGE_SIZE, nextLink);
        allUsers.push(...result.users);
        nextLink = result.nextLink;
        pageCount++;
        consecutiveErrors = 0; // Reset error counter on success

        // Update progress during fetch
        this.syncStatus.processedUsers = allUsers.length;
        this.syncStatus.progress = 0; // Fetching phase, will update to actual % later

        if (progressCallback && pageCount % 5 === 0) { // Update every 5 pages
          this.syncStatus.totalUsers = allUsers.length; // Temporary, will be final count later
          progressCallback(this.getSyncStatus());
        }

        console.log(`Fetched page ${pageCount}: ${allUsers.length} total users so far...`);

        // Small delay to avoid rate limiting
        await this.delay(300);

      } catch (error) {
        consecutiveErrors++;
        console.error(`Error fetching users page ${pageCount + 1} (attempt ${consecutiveErrors}/${MAX_CONSECUTIVE_ERRORS}):`, error);

        // If too many consecutive errors, stop
        if (consecutiveErrors >= MAX_CONSECUTIVE_ERRORS) {
          console.error('Too many consecutive errors. Stopping fetch.');
          // If we have some users, return them; otherwise throw
          if (allUsers.length === 0) {
            throw error;
          }
          break;
        }

        // Wait longer before retry after error
        await this.delay(2000);

        // If we have some users and hit an error, we can continue with what we have
        if (allUsers.length > 0) {
          console.log(`Continuing with ${allUsers.length} users already fetched`);
          break;
        }
      }
    } while (nextLink && !this.cancelRequested);

    console.log(`Fetch completed. Total users: ${allUsers.length} from ${pageCount} pages`);
    return allUsers;
  }

  /**
   * Process a batch of users
   * Split into smaller chunks to avoid overwhelming SharePoint
   */
  private async processBatch(users: IUserProfile[]): Promise<void> {
    const CHUNK_SIZE = 5; // Process 5 users at a time within each batch
    const CHUNK_DELAY = 500; // 500ms delay between chunks

    for (let i = 0; i < users.length; i += CHUNK_SIZE) {
      const chunk = users.slice(i, i + CHUNK_SIZE);
      const promises = chunk.map(user =>
        this.listService.addOrUpdateUser(user).catch(error => {
          console.error(`Error adding user ${user.userPrincipalName}:`, error);
          return null; // Don't fail entire batch
        })
      );

      await Promise.allSettled(promises);

      // Add delay between chunks (except after the last chunk)
      if (i + CHUNK_SIZE < users.length) {
        await this.delay(CHUNK_DELAY);
      }
    }
  }

  /**
   * Get default sync status
   */
  private getDefaultSyncStatus(): ISyncStatus {
    return {
      isRunning: false,
      totalUsers: 0,
      processedUsers: 0,
      currentBatch: 0,
      totalBatches: 0,
      errors: [],
      progress: 0
    };
  }

  /**
   * Delay helper
   */
  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Cancel ongoing sync (best effort)
   */
  public cancelSync(): void {
    if (this.syncStatus.isRunning) {
      console.log('Sync cancellation requested...');
      this.cancelRequested = true;
      // Note: Current batch will complete, but no new batches will start
    }
  }
}
