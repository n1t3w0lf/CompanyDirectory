import { ICacheEntry } from '../models/IUserProfile';
import { Constants } from '../models/Constants';

/**
 * IndexedDB helper for client-side caching
 */
export class CacheHelper {
  private dbName: string = Constants.INDEXEDDB_NAME;
  private dbVersion: number = Constants.INDEXEDDB_VERSION;
  private storeName: string = Constants.INDEXEDDB_STORE;
  private db: IDBDatabase | null = null;

  /**
   * Initialize IndexedDB
   */
  public async initialize(): Promise<void> {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(this.dbName, this.dbVersion);

      request.onerror = () => reject(request.error);
      request.onsuccess = () => {
        this.db = request.result;
        resolve();
      };

      request.onupgradeneeded = (event: IDBVersionChangeEvent) => {
        const db = (event.target as IDBOpenDBRequest).result;

        if (!db.objectStoreNames.contains(this.storeName)) {
          const objectStore = db.createObjectStore(this.storeName, { keyPath: 'key' });
          objectStore.createIndex('timestamp', 'timestamp', { unique: false });
        }
      };
    });
  }

  /**
   * Get item from cache
   */
  public async get<T>(key: string): Promise<T | null> {
    if (!this.db) await this.initialize();

    return new Promise((resolve, reject) => {
      const transaction = this.db!.transaction([this.storeName], 'readonly');
      const objectStore = transaction.objectStore(this.storeName);
      const request = objectStore.get(key);

      request.onerror = () => reject(request.error);
      request.onsuccess = () => {
        const result = request.result as ICacheEntry<T> | undefined;

        if (!result) {
          resolve(null);
          return;
        }

        // Check if cache entry is expired
        const now = Date.now();
        if (now - result.timestamp > result.ttl) {
          this.delete(key).catch(console.error);
          resolve(null);
          return;
        }

        resolve(result.data);
      };
    });
  }

  /**
   * Set item in cache
   */
  public async set<T>(key: string, data: T, ttl: number = Constants.CLIENT_CACHE_TTL): Promise<void> {
    if (!this.db) await this.initialize();

    return new Promise((resolve, reject) => {
      const transaction = this.db!.transaction([this.storeName], 'readwrite');
      const objectStore = transaction.objectStore(this.storeName);

      const cacheEntry: ICacheEntry<T> & { key: string } = {
        key,
        data,
        timestamp: Date.now(),
        ttl
      };

      const request = objectStore.put(cacheEntry);

      request.onerror = () => reject(request.error);
      request.onsuccess = () => resolve();
    });
  }

  /**
   * Delete item from cache
   */
  public async delete(key: string): Promise<void> {
    if (!this.db) await this.initialize();

    return new Promise((resolve, reject) => {
      const transaction = this.db!.transaction([this.storeName], 'readwrite');
      const objectStore = transaction.objectStore(this.storeName);
      const request = objectStore.delete(key);

      request.onerror = () => reject(request.error);
      request.onsuccess = () => resolve();
    });
  }

  /**
   * Clear all cache entries
   */
  public async clear(): Promise<void> {
    if (!this.db) await this.initialize();

    return new Promise((resolve, reject) => {
      const transaction = this.db!.transaction([this.storeName], 'readwrite');
      const objectStore = transaction.objectStore(this.storeName);
      const request = objectStore.clear();

      request.onerror = () => reject(request.error);
      request.onsuccess = () => resolve();
    });
  }

  /**
   * Get cache statistics
   */
  public async getStats(): Promise<{ count: number; size: number }> {
    if (!this.db) await this.initialize();

    return new Promise((resolve, reject) => {
      const transaction = this.db!.transaction([this.storeName], 'readonly');
      const objectStore = transaction.objectStore(this.storeName);
      const countRequest = objectStore.count();

      countRequest.onerror = () => reject(countRequest.error);
      countRequest.onsuccess = () => {
        resolve({
          count: countRequest.result,
          size: 0 // Size calculation would require iterating all entries
        });
      };
    });
  }
}

export const cacheHelper = new CacheHelper();
