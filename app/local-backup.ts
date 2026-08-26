export type StoredAppState = {
  slots: unknown[];
  activityLogs: unknown[];
  updatedAt: string;
};

export type SnapshotSource = 'automatic' | 'manual' | 'before-restore' | 'imported' | 'restored';

export type BackupSnapshot = StoredAppState & {
  id: string;
  source: SnapshotSource;
};

export type PortableBackup = {
  format: 'LPY-OT-CASE-BACKUP';
  schemaVersion: 1;
  appVersion: '16';
  createdAt: string;
  slots: unknown[];
  activityLogs: unknown[];
  checksum: string;
};

const DATABASE_NAME = 'lpy-ot-case-index';
const DATABASE_VERSION = 1;
const CURRENT_STATE_KEY = 'current';
const MAX_RECENT_SNAPSHOTS = 20;
const MAX_DAILY_AGE_DAYS = 30;

let databasePromise: Promise<IDBDatabase> | null = null;

function requestResult<T>(request: IDBRequest<T>) {
  return new Promise<T>((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error('IndexedDB request failed'));
  });
}

function transactionComplete(transaction: IDBTransaction) {
  return new Promise<void>((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onabort = () => reject(transaction.error || new Error('IndexedDB transaction aborted'));
    transaction.onerror = () => reject(transaction.error || new Error('IndexedDB transaction failed'));
  });
}

function openDatabase() {
  if (databasePromise) return databasePromise;
  databasePromise = new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
    request.onupgradeneeded = () => {
      const database = request.result;
      if (!database.objectStoreNames.contains('state')) database.createObjectStore('state', { keyPath: 'key' });
      if (!database.objectStoreNames.contains('snapshots')) {
        const store = database.createObjectStore('snapshots', { keyPath: 'id' });
        store.createIndex('createdAt', 'updatedAt');
      }
      if (!database.objectStoreNames.contains('metadata')) database.createObjectStore('metadata', { keyPath: 'key' });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error('未能開啟 IndexedDB'));
    request.onblocked = () => reject(new Error('IndexedDB 更新被其他視窗阻擋'));
  });
  return databasePromise;
}

export function supportsIndexedDB() {
  return typeof indexedDB !== 'undefined';
}

export async function loadCurrentState(): Promise<StoredAppState | null> {
  const database = await openDatabase();
  const transaction = database.transaction('state', 'readonly');
  const completed = transactionComplete(transaction);
  const result = await requestResult(transaction.objectStore('state').get(CURRENT_STATE_KEY));
  await completed;
  if (!result || typeof result !== 'object') return null;
  const record = result as StoredAppState & { key: string };
  return { slots: record.slots, activityLogs: record.activityLogs, updatedAt: record.updatedAt };
}

function makeSnapshot(state: StoredAppState, source: SnapshotSource): BackupSnapshot {
  return {
    ...state,
    id: `${state.updatedAt}-${Math.random().toString(36).slice(2, 9)}`,
    source,
  };
}

export async function saveCurrentState(state: StoredAppState, snapshotSource?: SnapshotSource) {
  const database = await openDatabase();
  const stores = snapshotSource ? ['state', 'snapshots'] : ['state'];
  const transaction = database.transaction(stores, 'readwrite');
  const completed = transactionComplete(transaction);
  transaction.objectStore('state').put({ key: CURRENT_STATE_KEY, ...state });
  if (snapshotSource) transaction.objectStore('snapshots').put(makeSnapshot(state, snapshotSource));
  await completed;
  if (snapshotSource) await pruneSnapshots();
}

export async function createLocalSnapshot(state: StoredAppState, source: SnapshotSource = 'manual') {
  const database = await openDatabase();
  const transaction = database.transaction('snapshots', 'readwrite');
  const completed = transactionComplete(transaction);
  transaction.objectStore('snapshots').put(makeSnapshot(state, source));
  await completed;
  await pruneSnapshots();
}

export async function listLocalSnapshots(): Promise<BackupSnapshot[]> {
  const database = await openDatabase();
  const transaction = database.transaction('snapshots', 'readonly');
  const completed = transactionComplete(transaction);
  const records = await requestResult(transaction.objectStore('snapshots').getAll()) as BackupSnapshot[];
  await completed;
  return records.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

async function pruneSnapshots() {
  const records = await listLocalSnapshots();
  const keep = new Set(records.slice(0, MAX_RECENT_SNAPSHOTS).map((record) => record.id));
  const daily = new Set<string>();
  const oldestAllowed = Date.now() - MAX_DAILY_AGE_DAYS * 24 * 60 * 60 * 1000;

  for (const record of records.slice(MAX_RECENT_SNAPSHOTS)) {
    const timestamp = Date.parse(record.updatedAt);
    const day = record.updatedAt.slice(0, 10);
    if (Number.isFinite(timestamp) && timestamp >= oldestAllowed && !daily.has(day)) {
      daily.add(day);
      keep.add(record.id);
    }
  }

  const removals = records.filter((record) => !keep.has(record.id));
  if (!removals.length) return;
  const database = await openDatabase();
  const transaction = database.transaction('snapshots', 'readwrite');
  const completed = transactionComplete(transaction);
  const store = transaction.objectStore('snapshots');
  removals.forEach((record) => store.delete(record.id));
  await completed;
}

export async function readMetadata(key: string): Promise<string | null> {
  const database = await openDatabase();
  const transaction = database.transaction('metadata', 'readonly');
  const completed = transactionComplete(transaction);
  const result = await requestResult(transaction.objectStore('metadata').get(key));
  await completed;
  return result && typeof result === 'object' && typeof (result as { value?: unknown }).value === 'string'
    ? (result as { value: string }).value : null;
}

export async function writeMetadata(key: string, value: string) {
  const database = await openDatabase();
  const transaction = database.transaction('metadata', 'readwrite');
  const completed = transactionComplete(transaction);
  transaction.objectStore('metadata').put({ key, value });
  await completed;
}

async function sha256(value: string) {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(value));
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

export async function createPortableBackup(slots: unknown[], activityLogs: unknown[]): Promise<PortableBackup> {
  const payload = {
    format: 'LPY-OT-CASE-BACKUP' as const,
    schemaVersion: 1 as const,
    appVersion: '16' as const,
    createdAt: new Date().toISOString(),
    slots,
    activityLogs,
  };
  return { ...payload, checksum: await sha256(JSON.stringify(payload)) };
}

export async function hasValidChecksum(backup: PortableBackup) {
  const { checksum, ...payload } = backup;
  return checksum === await sha256(JSON.stringify(payload));
}
