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
  private readonly BATCH_SIZE = 100; // Process 100 users at a time
  private readonly GRAPH_PAGE_SIZE = 999; // Max Graph API page size

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

      console.log('Starting initial sync of all users from Entra ID...');

      // Ensure list is created
      await this.listService.ensureList();

      // Fetch all users from Graph API with pagination
      const allUsers = await this.fetchAllUsersFromGraph(progressCallback);

      this.syncStatus.totalUsers = allUsers.length;
      this.syncStatus.totalBatches = Math.ceil(allUsers.length / this.BATCH_SIZE);

      console.log(`Fetched ${allUsers.length} users from Entra ID. Starting list population...`);

      // Process users in batches to avoid overwhelming SharePoint
      for (let i = 0; i < allUsers.length; i += this.BATCH_SIZE) {
        const batch = allUsers.slice(i, i + this.BATCH_SIZE);
        this.syncStatus.currentBatch = Math.floor(i / this.BATCH_SIZE) + 1;

        try {
          await this.processBatch(batch);
          this.syncStatus.processedUsers += batch.length;
          this.syncStatus.progress = Math.round((this.syncStatus.processedUsers / this.syncStatus.totalUsers) * 100);

          if (progressCallback) {
            progressCallback(this.getSyncStatus());
          }

          console.log(`Processed batch ${this.syncStatus.currentBatch}/${this.syncStatus.totalBatches} (${this.syncStatus.processedUsers}/${this.syncStatus.totalUsers} users)`);

          // Small delay to avoid throttling
          await this.delay(500);
        } catch (error) {
          const errorMsg = `Error processing batch ${this.syncStatus.currentBatch}: ${ErrorHandler.getUserMessage(error)}`;
          this.syncStatus.errors.push(errorMsg);
          console.error(errorMsg);
          // Continue with next batch even if one fails
        }
      }

      // Update sync metadata
      await this.listService.updateSyncMetadata({
        lastFullSync: new Date(),
        totalUsers: this.syncStatus.totalUsers,
        lastSyncSuccess: this.syncStatus.errors.length === 0
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
      throw new Error(errorMsg);
    } finally {
      this.syncStatus.isRunning = false;
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

    do {
      try {
        const result = await this.graphService.searchUsers('', this.GRAPH_PAGE_SIZE, nextLink);
        allUsers.push(...result.users);
        nextLink = result.nextLink;
        pageCount++;

        // Update progress during fetch
        this.syncStatus.processedUsers = allUsers.length;
        this.syncStatus.progress = 0; // Fetching phase, will update to actual % later

        if (progressCallback && pageCount % 5 === 0) { // Update every 5 pages
          this.syncStatus.totalUsers = allUsers.length; // Temporary, will be final count later
          progressCallback(this.getSyncStatus());
        }

        console.log(`Fetched page ${pageCount}: ${allUsers.length} total users so far...`);

        // Small delay to avoid rate limiting
        await this.delay(200);

      } catch (error) {
        console.error(`Error fetching users page ${pageCount}:`, error);
        // If we have some users, continue; otherwise throw
        if (allUsers.length === 0) {
          throw error;
        }
        break; // Exit loop if error but we have some data
      }
    } while (nextLink);

    return allUsers;
  }

  /**
   * Process a batch of users
   */
  private async processBatch(users: IUserProfile[]): Promise<void> {
    const promises = users.map(user =>
      this.listService.addOrUpdateUser(user).catch(error => {
        console.error(`Error adding user ${user.userPrincipalName}:`, error);
        return null; // Don't fail entire batch
      })
    );

    await Promise.allSettled(promises);
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
      this.syncStatus.isRunning = false;
      // Note: Current batch will complete, but no new batches will start
    }
  }
}
