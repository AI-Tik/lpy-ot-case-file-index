'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import { exportExcel } from './excel-export';
import {
  createLocalSnapshot,
  createPortableBackup,
  hasValidChecksum,
  listLocalSnapshots,
  loadCurrentState,
  readMetadata,
  saveCurrentState,
  supportsIndexedDB,
  writeMetadata,
  type BackupSnapshot,
  type PortableBackup,
} from './local-backup';

type FileColor = 'blue' | 'red' | 'orange';
type FlagKey = 'c' | 'h' | 'e' | 'd' | 'v';
type Slot = { color: FileColor; number: number; name: string; c: boolean; h: boolean; e: boolean; d: boolean; v: boolean };
type SheetMode = 'actions' | 'edit' | 'delete';
type AppView = 'search' | 'files' | 'batch' | 'history' | 'backup';
type BatchMode = 'add' | 'delete';
type BatchAddRow = { id: number; name: string; position: string; c: boolean; h: boolean; e: boolean; d: boolean; v: boolean };
type ActivityLog = {
  id: string;
  type: 'add' | 'delete';
  timestamp: string;
  name: string;
  color: FileColor;
  number: number;
  flags: string[];
};

type RestoreCandidate = {
  label: string;
  createdAt: string;
  slots: Slot[];
  activityLogs: ActivityLog[];
  source: 'file' | 'snapshot';
  added: number;
  removed: number;
  modified: number;
};

type SpeechRecognitionLike = {
  lang: string;
  continuous: boolean;
  interimResults: boolean;
  maxAlternatives: number;
  phrases?: unknown[];
  start: () => void;
  stop: () => void;
  onresult: ((event: { results: { 0: { length: number; [key: number]: { transcript: string } } } }) => void) | null;
  onerror: (() => void) | null;
  onend: (() => void) | null;
};

type SpeechWindow = Window & {
  SpeechRecognition?: new () => SpeechRecognitionLike;
  webkitSpeechRecognition?: new () => SpeechRecognitionLike;
  SpeechRecognitionPhrase?: new (phrase: string, boost?: number) => unknown;
};

const COLOR_INFO: Record<FileColor, { label: string; short: string }> = {
  blue: { label: '藍色', short: '藍' },
  red: { label: '紅色', short: '紅' },
  orange: { label: '橙色', short: '橙' },
};

const PUBLIC_BASE_PATH = process.env.NEXT_PUBLIC_BASE_PATH ?? '';

const FLAG_ORDER: Array<{ key: FlagKey; label: string }> = [
  { key: 'c', label: 'C' },
  { key: 'h', label: 'H' },
  { key: 'e', label: 'E' },
  { key: 'd', label: 'D' },
  { key: 'v', label: 'V' },
];

const SLOT_COUNT = 31;

const INITIAL_NAMES: Record<FileColor, string[]> = {
  blue: Array.from({ length: SLOT_COUNT }, () => ''),
  red: Array.from({ length: SLOT_COUNT }, () => ''),
  orange: Array.from({ length: SLOT_COUNT }, () => ''),
};

const INITIAL_FLAGS: Record<FileColor, { d: number[]; v: number[] }> = {
  blue: { d: [], v: [] },
  red: { d: [], v: [] },
  orange: { d: [], v: [] },
};

const COLORS: FileColor[] = ['blue', 'red', 'orange'];
const DISPLAY_COLORS: FileColor[] = ['red', 'orange', 'blue'];
const STORAGE_KEY = 'lpy-ot-case-file-index-v3';
const PREVIOUS_STORAGE_KEY = 'lpy-ot-case-file-index-v2';
const OLD_STORAGE_KEY = 'lpy-ot-case-file-index-v1';
const ACTIVITY_STORAGE_KEY = 'lpy-ot-case-file-activity-v1';
const LAST_EXTERNAL_BACKUP_KEY = 'lastExternalBackupAt';
const LOCAL_UPDATED_AT_KEY = 'lpy-ot-case-file-updated-at-v1';
const DEFAULT_SLOTS: Slot[] = COLORS.flatMap((color) =>
  INITIAL_NAMES[color].map((name, index) => {
    const number = index + 1;
    return {
      color,
      number,
      name,
      c: false,
      h: false,
      e: false,
      d: INITIAL_FLAGS[color].d.includes(number),
      v: INITIAL_FLAGS[color].v.includes(number),
    };
  }),
);

const slotKey = (slot: Pick<Slot, 'color' | 'number'>) => `${slot.color}-${slot.number}`;
const emptyBatchRow = (id: number): BatchAddRow => ({
  id, name: '', position: 'random', c: false, h: false, e: false, d: false, v: false,
});
const activeFlags = (slot: Pick<Slot, FlagKey>) => FLAG_ORDER.filter((flag) => slot[flag.key]).map((flag) => flag.label);
const formatLogDate = (value: string) => new Intl.DateTimeFormat('zh-HK', {
  year: '2-digit', month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit', hour12: false,
  timeZone: 'Asia/Hong_Kong',
}).format(new Date(value));
const formatBackupDate = (value: string) => new Intl.DateTimeFormat('zh-HK', {
  year: 'numeric', month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit', hour12: false,
  timeZone: 'Asia/Hong_Kong',
}).format(new Date(value));

function isValidSlots(value: unknown): value is Slot[] {
  if (!Array.isArray(value) || value.length !== 93) return false;
  const validItems = value.every((item) => item && typeof item === 'object'
    && COLORS.includes((item as Slot).color)
    && Number.isInteger((item as Slot).number) && (item as Slot).number >= 1 && (item as Slot).number <= 31
    && typeof (item as Slot).name === 'string'
    && typeof (item as Slot).c === 'boolean'
    && typeof (item as Slot).h === 'boolean'
    && typeof (item as Slot).e === 'boolean'
    && typeof (item as Slot).d === 'boolean'
    && typeof (item as Slot).v === 'boolean');
  if (!validItems) return false;
  return new Set(value.map((slot) => slotKey(slot as Slot))).size === 93;
}

function normalizeSlots(value: Slot[]) {
  const byPosition = new Map(value.map((slot) => [slotKey(slot), slot]));
  return DEFAULT_SLOTS.map((slot) => ({ ...byPosition.get(slotKey(slot))! }));
}

function isPreviousSlots(value: unknown): value is Array<Omit<Slot, 'c' | 'h' | 'e'>> {
  return Array.isArray(value) && value.length === 93
    && value.every((item) => item && typeof item === 'object'
      && typeof (item as Slot).name === 'string'
      && typeof (item as Slot).d === 'boolean'
      && typeof (item as Slot).v === 'boolean');
}

function isOldSlots(value: unknown): value is Array<Pick<Slot, 'color' | 'number' | 'name'>> {
  return Array.isArray(value) && value.length === 93
    && value.every((item) => item && typeof item === 'object' && typeof item.name === 'string');
}

function isActivityLogs(value: unknown): value is ActivityLog[] {
  return Array.isArray(value) && value.every((item) => item && typeof item === 'object'
    && ['add', 'delete'].includes((item as ActivityLog).type)
    && typeof (item as ActivityLog).id === 'string'
    && typeof (item as ActivityLog).timestamp === 'string' && !Number.isNaN(Date.parse((item as ActivityLog).timestamp))
    && typeof (item as ActivityLog).name === 'string'
    && COLORS.includes((item as ActivityLog).color)
    && Number.isInteger((item as ActivityLog).number) && (item as ActivityLog).number >= 1 && (item as ActivityLog).number <= 31
    && Array.isArray((item as ActivityLog).flags)
    && (item as ActivityLog).flags.every((flag) => FLAG_ORDER.some((allowed) => allowed.label === flag)));
}

function isPortableBackup(value: unknown): value is PortableBackup & { slots: Slot[]; activityLogs: ActivityLog[] } {
  if (!value || typeof value !== 'object') return false;
  const backup = value as PortableBackup;
  return backup.format === 'LPY-OT-CASE-BACKUP'
    && backup.schemaVersion === 1
    && typeof backup.appVersion === 'string'
    && typeof backup.createdAt === 'string'
    && !Number.isNaN(Date.parse(backup.createdAt))
    && typeof backup.checksum === 'string'
    && isValidSlots(backup.slots)
    && isActivityLogs(backup.activityLogs);
}

function legacyStoredState() {
  let nextSlots = DEFAULT_SLOTS;
  let nextActivityLogs: ActivityLog[] = [];
  try {
    const current = localStorage.getItem(STORAGE_KEY);
    const parsed: unknown = current ? JSON.parse(current) : null;
    if (isValidSlots(parsed)) {
      nextSlots = normalizeSlots(parsed);
    } else {
      const previous = localStorage.getItem(PREVIOUS_STORAGE_KEY);
      const previousParsed: unknown = previous ? JSON.parse(previous) : null;
      if (isPreviousSlots(previousParsed)) {
        nextSlots = previousParsed.map((slot) => ({ ...slot, c: false, h: false, e: false }));
      } else {
        const old = localStorage.getItem(OLD_STORAGE_KEY);
        const oldParsed: unknown = old ? JSON.parse(old) : null;
        if (isOldSlots(oldParsed)) {
          nextSlots = DEFAULT_SLOTS.map((slot, index) => ({ ...slot, name: oldParsed[index].name }));
        }
      }
    }
    const storedActivity = localStorage.getItem(ACTIVITY_STORAGE_KEY);
    const parsedActivity: unknown = storedActivity ? JSON.parse(storedActivity) : null;
    if (isActivityLogs(parsedActivity)) {
      nextActivityLogs = [...parsedActivity].sort((a, b) => b.timestamp.localeCompare(a.timestamp));
    }
  } catch {
    // Keep the Excel baseline if legacy storage cannot be read.
  }
  return { slots: nextSlots, activityLogs: nextActivityLogs };
}

function editDistance(a: string, b: string) {
  const matrix = Array.from({ length: a.length + 1 }, () => Array<number>(b.length + 1).fill(0));
  for (let i = 0; i <= a.length; i += 1) matrix[i][0] = i;
  for (let j = 0; j <= b.length; j += 1) matrix[0][j] = j;
  for (let i = 1; i <= a.length; i += 1) {
    for (let j = 1; j <= b.length; j += 1) {
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1),
      );
    }
  }
  return matrix[a.length][b.length];
}

function FlagBadges({ slot }: { slot: Pick<Slot, FlagKey> }) {
  const flags = activeFlags(slot);
  if (!flags.length) return null;
  return <span className="flag-badges">{flags.map((flag) => <b key={flag}>{flag}</b>)}</span>;
}

function LogFlagBadges({ flags }: { flags: string[] }) {
  if (!flags.length) return <span className="no-flags">—</span>;
  return <span className="flag-badges log-flags">{flags.map((flag) => <b key={flag}>{flag}</b>)}</span>;
}

function LongPressSlot({ slot, onAdd, onManage }: { slot: Slot; onAdd: () => void; onManage: () => void }) {
  const [pressing, setPressing] = useState(false);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const startPoint = useRef({ x: 0, y: 0 });
  const completed = useRef(false);

  const cancel = () => {
    if (timer.current) clearTimeout(timer.current);
    timer.current = null;
    setPressing(false);
  };

  const begin = (event: React.PointerEvent<HTMLButtonElement>) => {
    if (!slot.name) return;
    completed.current = false;
    startPoint.current = { x: event.clientX, y: event.clientY };
    setPressing(true);
    timer.current = setTimeout(() => {
      completed.current = true;
      setPressing(false);
      if (navigator.vibrate) navigator.vibrate(45);
      onManage();
    }, 3000);
  };

  const move = (event: React.PointerEvent<HTMLButtonElement>) => {
    if (Math.abs(event.clientX - startPoint.current.x) > 12
      || Math.abs(event.clientY - startPoint.current.y) > 12) cancel();
  };

  useEffect(() => cancel, []);

  return (
    <button
      type="button"
      className={`slot-card ${slot.color} ${slot.name ? '' : 'empty'} ${pressing ? 'pressing' : ''}`}
      onPointerDown={begin}
      onPointerMove={move}
      onPointerUp={cancel}
      onPointerCancel={cancel}
      onPointerLeave={cancel}
      onContextMenu={(event) => event.preventDefault()}
      onClick={() => {
        if (completed.current) { completed.current = false; return; }
        if (!slot.name) onAdd();
      }}
      aria-label={slot.name ? `${slot.name}，長按三秒管理` : `${COLOR_INFO[slot.color].label}${slot.number}號，加入個案`}
    >
      <span className="slot-number">{slot.number}</span>
      <span className="slot-name">{slot.name || '＋ 空位'}</span>
      <FlagBadges slot={slot} />
      {slot.name && <span className="hold-progress" aria-hidden="true" />}
    </button>
  );
}

export default function Home() {
  const [slots, setSlots] = useState<Slot[]>(DEFAULT_SLOTS);
  const [dataReady, setDataReady] = useState(false);
  const [view, setView] = useState<AppView>('search');
  const [query, setQuery] = useState('');
  const [suggestionsOpen, setSuggestionsOpen] = useState(false);
  const [listening, setListening] = useState(false);
  const [speechNote, setSpeechNote] = useState('');
  const [exporting, setExporting] = useState(false);
  const [selected, setSelected] = useState<Slot | null>(null);
  const [sheetMode, setSheetMode] = useState<SheetMode>('actions');
  const [draftName, setDraftName] = useState('');
  const [draftC, setDraftC] = useState(false);
  const [draftH, setDraftH] = useState(false);
  const [draftE, setDraftE] = useState(false);
  const [draftD, setDraftD] = useState(false);
  const [draftV, setDraftV] = useState(false);
  const [error, setError] = useState('');
  const [deleteCode, setDeleteCode] = useState('');
  const [toast, setToast] = useState('');
  const [batchMode, setBatchMode] = useState<BatchMode>('add');
  const [batchRows, setBatchRows] = useState<BatchAddRow[]>([
    emptyBatchRow(1),
    emptyBatchRow(2),
  ]);
  const [batchAddError, setBatchAddError] = useState('');
  const [batchDeleteSearch, setBatchDeleteSearch] = useState('');
  const [batchSelected, setBatchSelected] = useState<string[]>([]);
  const [batchDeleteOpen, setBatchDeleteOpen] = useState(false);
  const [batchDeleteCode, setBatchDeleteCode] = useState('');
  const [activityLogs, setActivityLogs] = useState<ActivityLog[]>([]);
  const [storageMode, setStorageMode] = useState<'loading' | 'indexeddb' | 'fallback'>('loading');
  const [storageSaving, setStorageSaving] = useState(false);
  const [lastSavedAt, setLastSavedAt] = useState('');
  const [lastExternalBackupAt, setLastExternalBackupAt] = useState('');
  const [snapshots, setSnapshots] = useState<BackupSnapshot[]>([]);
  const [backupWorking, setBackupWorking] = useState(false);
  const [backupError, setBackupError] = useState('');
  const [pendingRestore, setPendingRestore] = useState<RestoreCandidate | null>(null);
  const [restoreCode, setRestoreCode] = useState('');
  const [restoreWorking, setRestoreWorking] = useState(false);
  const searchRef = useRef<HTMLInputElement>(null);
  const recognitionRef = useRef<SpeechRecognitionLike | null>(null);
  const nextBatchRowId = useRef(3);
  const importBackupRef = useRef<HTMLInputElement>(null);
  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const lastStoredFingerprint = useRef('');

  useEffect(() => {
    let cancelled = false;
    const initializeStorage = async () => {
      const legacy = legacyStoredState();
      let nextSlots = legacy.slots;
      let nextActivityLogs = legacy.activityLogs;
      let savedAt = new Date().toISOString();
      let nextStorageMode: 'indexeddb' | 'fallback' = 'fallback';
      let nextSnapshots: BackupSnapshot[] = [];
      let externalBackupAt = localStorage.getItem(LAST_EXTERNAL_BACKUP_KEY) || '';
      const legacyUpdatedAt = localStorage.getItem(LOCAL_UPDATED_AT_KEY) || '';

      if (supportsIndexedDB()) {
        try {
          const stored = await loadCurrentState();
          if (stored && isValidSlots(stored.slots) && isActivityLogs(stored.activityLogs)
            && (!legacyUpdatedAt || stored.updatedAt >= legacyUpdatedAt)) {
            nextSlots = normalizeSlots(stored.slots);
            nextActivityLogs = [...stored.activityLogs].sort((a, b) => b.timestamp.localeCompare(a.timestamp));
            savedAt = stored.updatedAt;
          } else {
            savedAt = legacyUpdatedAt || new Date().toISOString();
            await saveCurrentState({ slots: nextSlots, activityLogs: nextActivityLogs, updatedAt: savedAt }, 'manual');
          }
          nextSnapshots = (await listLocalSnapshots())
            .filter((snapshot) => isValidSlots(snapshot.slots) && isActivityLogs(snapshot.activityLogs));
          externalBackupAt = await readMetadata(LAST_EXTERNAL_BACKUP_KEY)
            || localStorage.getItem(LAST_EXTERNAL_BACKUP_KEY) || '';
          nextStorageMode = 'indexeddb';
        } catch {
          nextStorageMode = 'fallback';
        }
      }

      if (cancelled) return;
      const normalizedSlots = normalizeSlots(nextSlots);
      const normalizedLogs = nextActivityLogs.slice(0, 1500);
      setSlots(normalizedSlots);
      setActivityLogs(normalizedLogs);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(normalizedSlots));
      localStorage.setItem(ACTIVITY_STORAGE_KEY, JSON.stringify(normalizedLogs));
      localStorage.setItem(LOCAL_UPDATED_AT_KEY, savedAt);
      lastStoredFingerprint.current = JSON.stringify([normalizedSlots, normalizedLogs]);
      setLastSavedAt(savedAt);
      setLastExternalBackupAt(externalBackupAt);
      setSnapshots(nextSnapshots);
      setStorageMode(nextStorageMode);
      setDataReady(true);
    };
    initializeStorage();
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!dataReady) return;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(slots));
    localStorage.setItem(ACTIVITY_STORAGE_KEY, JSON.stringify(activityLogs));
    const fingerprint = JSON.stringify([slots, activityLogs]);
    if (fingerprint === lastStoredFingerprint.current) return;
    const updatedAt = new Date().toISOString();
    localStorage.setItem(LOCAL_UPDATED_AT_KEY, updatedAt);
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    setStorageSaving(true);
    saveTimerRef.current = setTimeout(async () => {
      try {
        if (storageMode === 'indexeddb') {
          await saveCurrentState({ slots, activityLogs, updatedAt }, 'automatic');
          const nextSnapshots = await listLocalSnapshots();
          setSnapshots(nextSnapshots.filter((snapshot) => isValidSlots(snapshot.slots) && isActivityLogs(snapshot.activityLogs)));
        }
        lastStoredFingerprint.current = fingerprint;
        setLastSavedAt(updatedAt);
      } catch {
        setStorageMode('fallback');
      } finally {
        setStorageSaving(false);
      }
    }, 450);
    return () => {
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    };
  }, [activityLogs, dataReady, slots, storageMode]);

  useEffect(() => {
    if (!toast) return;
    const timer = setTimeout(() => setToast(''), 2400);
    return () => clearTimeout(timer);
  }, [toast]);

  useEffect(() => {
    if (!('serviceWorker' in navigator)) return;
    const registerOfflineApp = () => {
      navigator.serviceWorker.register(`${PUBLIC_BASE_PATH}/sw.js`)
        .then(() => navigator.storage?.persist?.())
        .catch(() => {
          // The online app remains usable if offline setup is unavailable.
        });
    };
    if (document.readyState === 'complete') registerOfflineApp();
    else window.addEventListener('load', registerOfflineApp, { once: true });
    return () => window.removeEventListener('load', registerOfflineApp);
  }, []);

  const names = useMemo(() => slots.filter((slot) => slot.name).map((slot) => slot.name), [slots]);
  const normalizedQuery = query.trim().toLocaleLowerCase('zh-Hant');
  const queryFlag = FLAG_ORDER.find((flag) => flag.label.toLocaleLowerCase() === normalizedQuery);
  const results = useMemo(() => {
    if (!normalizedQuery) return [];
    if (queryFlag) return slots.filter((slot) => slot.name && slot[queryFlag.key]);
    return slots.filter((slot) => slot.name.toLocaleLowerCase('zh-Hant').includes(normalizedQuery));
  }, [normalizedQuery, queryFlag, slots]);

  const suggestions = useMemo(() => {
    if (!normalizedQuery) return names.slice(0, 8);
    const flagSuggestions = FLAG_ORDER.map((flag) => flag.label)
      .filter((flag) => flag.toLocaleLowerCase().startsWith(normalizedQuery));
    const nameSuggestions = [...names]
      .map((name) => ({
        name,
        direct: name.toLocaleLowerCase('zh-Hant').includes(normalizedQuery),
        distance: editDistance(name.toLocaleLowerCase('zh-Hant'), normalizedQuery),
      }))
      .sort((a, b) => Number(b.direct) - Number(a.direct) || a.distance - b.distance)
      .slice(0, 8)
      .map((item) => item.name);
    return [...flagSuggestions, ...nameSuggestions].slice(0, 8);
  }, [names, normalizedQuery]);

  const occupied = slots.filter((slot) => slot.name).length;
  const emptySlots = useMemo(() => DISPLAY_COLORS.flatMap((color) =>
    slots.filter((slot) => slot.color === color && !slot.name)), [slots]);
  const occupiedSlots = useMemo(() => slots.filter((slot) => slot.name), [slots]);
  const filteredDeleteSlots = useMemo(() => {
    const clean = batchDeleteSearch.trim().toLocaleLowerCase('zh-Hant');
    return clean ? occupiedSlots.filter((slot) => slot.name.toLocaleLowerCase('zh-Hant').includes(clean)) : occupiedSlots;
  }, [batchDeleteSearch, occupiedSlots]);
  const addLogs = useMemo(() => activityLogs.filter((log) => log.type === 'add'), [activityLogs]);
  const deleteLogs = useMemo(() => activityLogs.filter((log) => log.type === 'delete'), [activityLogs]);

  const chooseSuggestion = (name: string) => {
    setQuery(name);
    setSuggestionsOpen(false);
    setSpeechNote('');
  };

  const appendActivity = (type: ActivityLog['type'], affectedSlots: Slot[]) => {
    const timestamp = new Date().toISOString();
    const entries: ActivityLog[] = affectedSlots.map((slot, index) => ({
      id: `${Date.now()}-${index}-${Math.random().toString(36).slice(2, 8)}`,
      type,
      timestamp,
      name: slot.name,
      color: slot.color,
      number: slot.number,
      flags: activeFlags(slot),
    }));
    setActivityLogs((current) => [...entries, ...current].slice(0, 1500));
  };

  const closestName = (alternatives: string[]) => {
    let best = { name: '', score: Number.POSITIVE_INFINITY };
    for (const alternative of alternatives) {
      const clean = alternative.replace(/[，。,.!?！？\s]/g, '');
      for (const name of names) {
        const distance = editDistance(clean, name);
        const score = clean === name ? 0
          : clean.includes(name) || name.includes(clean) ? 0.2
            : distance / Math.max(clean.length, name.length, 1);
        if (score < best.score) best = { name, score };
      }
    }
    return best.score <= 0.5 ? best.name : alternatives[0];
  };

  const startVoiceSearch = () => {
    if (listening) { recognitionRef.current?.stop(); return; }
    if (!navigator.onLine) {
      setSpeechNote('離線時請使用鍵盤輸入或從下拉名單選擇。');
      searchRef.current?.focus();
      return;
    }
    const speechWindow = window as SpeechWindow;
    const Recognition = speechWindow.SpeechRecognition || speechWindow.webkitSpeechRecognition;
    if (!Recognition) {
      setSpeechNote('此瀏覽器未支援網頁語音輸入，可使用手機鍵盤的咪高峰。');
      searchRef.current?.focus();
      return;
    }
    const recognition = new Recognition();
    recognitionRef.current = recognition;
    recognition.lang = 'zh-HK';
    recognition.continuous = false;
    recognition.interimResults = false;
    recognition.maxAlternatives = 10;
    if (speechWindow.SpeechRecognitionPhrase) {
      recognition.phrases = names.map((name) => new speechWindow.SpeechRecognitionPhrase!(name, 8));
    }
    recognition.onresult = (event) => {
      const firstResult = event.results[0];
      const alternatives = Array.from({ length: firstResult.length }, (_, index) => firstResult[index].transcript);
      const matched = closestName(alternatives);
      setQuery(matched);
      setSuggestionsOpen(true);
      setSpeechNote(matched && names.includes(matched) ? `已按名單配對為「${matched}」` : '請從下拉名單確認姓名');
    };
    recognition.onerror = () => setSpeechNote('未能辨識，請再試一次或使用手機鍵盤咪高峰。');
    recognition.onend = () => setListening(false);
    setListening(true);
    setSpeechNote('正在聆聽，請讀出完整姓名…');
    recognition.start();
  };

  const openEditor = (slot: Slot) => {
    setSelected(slot);
    setDraftName(slot.name);
    setDraftC(slot.c);
    setDraftH(slot.h);
    setDraftE(slot.e);
    setDraftD(slot.d);
    setDraftV(slot.v);
    setError('');
    setSheetMode('edit');
  };

  const openActions = (slot: Slot) => {
    setSelected(slot);
    setDraftName(slot.name);
    setDraftC(slot.c);
    setDraftH(slot.h);
    setDraftE(slot.e);
    setDraftD(slot.d);
    setDraftV(slot.v);
    setError('');
    setSheetMode('actions');
  };

  const closeSheet = () => {
    setSelected(null);
    setDraftName('');
    setError('');
    setDeleteCode('');
  };

  const saveName = () => {
    if (!selected) return;
    const cleanName = draftName.trim();
    if (!cleanName) { setError('請輸入個案名稱'); return; }
    const duplicate = slots.find((slot) => slot.name === cleanName
      && !(slot.color === selected.color && slot.number === selected.number));
    if (duplicate) {
      setError(`這個名稱已在${COLOR_INFO[duplicate.color].label} ${duplicate.number} 號`);
      return;
    }
    const updatedSlot = { ...selected, name: cleanName, c: draftC, h: draftH, e: draftE, d: draftD, v: draftV };
    setSlots((current) => current.map((slot) => slot.color === selected.color && slot.number === selected.number
      ? updatedSlot : slot));
    if (!selected.name) appendActivity('add', [updatedSlot]);
    setToast(selected.name ? '資料已更新' : '新個案已加入');
    closeSheet();
  };

  const deleteName = () => {
    if (!selected || deleteCode !== 'LPY') return;
    appendActivity('delete', [selected]);
    setSlots((current) => current.map((slot) => slot.color === selected.color && slot.number === selected.number
      ? { ...slot, name: '', c: false, h: false, e: false, d: false, v: false } : slot));
    setToast('個案已刪除，位置現已空置');
    closeSheet();
  };

  const updateBatchRow = (id: number, changes: Partial<BatchAddRow>) => {
    setBatchRows((current) => current.map((row) => row.id === id ? { ...row, ...changes } : row));
    setBatchAddError('');
  };

  const addBatchRow = () => {
    const id = nextBatchRowId.current;
    nextBatchRowId.current += 1;
    setBatchRows((current) => [...current, emptyBatchRow(id)]);
  };

  const removeBatchRow = (id: number) => {
    setBatchRows((current) => current.length === 1
      ? [emptyBatchRow(current[0].id)]
      : current.filter((row) => row.id !== id));
    setBatchAddError('');
  };

  const addMultipleCases = () => {
    const rows = batchRows.map((row) => ({ ...row, name: row.name.trim() })).filter((row) => row.name);
    if (!rows.length) { setBatchAddError('請至少輸入一個個案名稱'); return; }
    if (rows.length > emptySlots.length) { setBatchAddError(`只剩下 ${emptySlots.length} 個空位`); return; }

    const submittedNames = rows.map((row) => row.name);
    const duplicateSubmitted = submittedNames.find((name, index) => submittedNames.indexOf(name) !== index);
    if (duplicateSubmitted) { setBatchAddError(`「${duplicateSubmitted}」在新增名單中重複`); return; }
    const duplicateExisting = rows.find((row) => slots.some((slot) => slot.name === row.name));
    if (duplicateExisting) { setBatchAddError(`「${duplicateExisting.name}」已在個案名單內`); return; }

    const chosenPositions = rows.filter((row) => row.position !== 'random').map((row) => row.position);
    const duplicatePosition = chosenPositions.find((key, index) => chosenPositions.indexOf(key) !== index);
    if (duplicatePosition) { setBatchAddError('同一個空位不能分配給兩位個案'); return; }

    const available = [...emptySlots];
    const assignments = new Map<number, string>();
    for (const row of rows.filter((item) => item.position !== 'random')) {
      const index = available.findIndex((slot) => slotKey(slot) === row.position);
      if (index < 0) { setBatchAddError('選取的空位已被使用，請重新選擇'); return; }
      assignments.set(row.id, row.position);
      available.splice(index, 1);
    }
    for (const row of rows.filter((item) => item.position === 'random')) {
      const index = Math.floor(Math.random() * available.length);
      const assigned = available.splice(index, 1)[0];
      assignments.set(row.id, slotKey(assigned));
    }

    const additions = rows.map((row) => {
      const position = assignments.get(row.id)!;
      const slot = emptySlots.find((item) => slotKey(item) === position)!;
      return { ...slot, name: row.name, c: row.c, h: row.h, e: row.e, d: row.d, v: row.v };
    });
    const additionsByPosition = new Map(additions.map((slot) => [slotKey(slot), slot]));
    setSlots((current) => current.map((slot) => {
      const addition = additionsByPosition.get(slotKey(slot));
      return addition || slot;
    }));
    appendActivity('add', additions);
    setBatchRows([
      emptyBatchRow(nextBatchRowId.current++),
      emptyBatchRow(nextBatchRowId.current++),
    ]);
    setBatchAddError('');
    setToast(`已新增 ${rows.length} 位個案並分配位置`);
  };

  const toggleBatchSelection = (key: string) => {
    setBatchSelected((current) => current.includes(key)
      ? current.filter((item) => item !== key)
      : [...current, key]);
  };

  const deleteMultipleCases = () => {
    if (batchDeleteCode !== 'LPY' || !batchSelected.length) return;
    const selectedKeys = new Set(batchSelected);
    const count = batchSelected.length;
    const removedSlots = slots.filter((slot) => selectedKeys.has(slotKey(slot)));
    appendActivity('delete', removedSlots);
    setSlots((current) => current.map((slot) => selectedKeys.has(slotKey(slot))
      ? { ...slot, name: '', c: false, h: false, e: false, d: false, v: false } : slot));
    setBatchSelected([]);
    setBatchDeleteCode('');
    setBatchDeleteOpen(false);
    setToast(`已刪除 ${count} 位個案，位置現已空置`);
  };

  const downloadExcel = async () => {
    setExporting(true);
    try {
      await exportExcel(slots);
      setToast('Excel 名單已輸出');
    } catch {
      setToast('未能輸出 Excel，請稍後再試');
    } finally {
      setExporting(false);
    }
  };

  const refreshSnapshots = async () => {
    if (storageMode !== 'indexeddb') return;
    const nextSnapshots = await listLocalSnapshots();
    setSnapshots(nextSnapshots.filter((snapshot) => isValidSlots(snapshot.slots) && isActivityLogs(snapshot.activityLogs)));
  };

  const saveManualLocalVersion = async () => {
    if (storageMode !== 'indexeddb') {
      setBackupError('這個瀏覽器暫時未能使用加強本機版本；請建立完整備份檔。');
      return;
    }
    setBackupWorking(true);
    setBackupError('');
    try {
      await createLocalSnapshot({ slots, activityLogs, updatedAt: new Date().toISOString() }, 'manual');
      await refreshSnapshots();
      setToast('已保存本機版本');
    } catch {
      setBackupError('未能保存本機版本，請改為下載完整備份。');
    } finally {
      setBackupWorking(false);
    }
  };

  const downloadFullBackup = async () => {
    setBackupWorking(true);
    setBackupError('');
    try {
      const backup = await createPortableBackup(slots, activityLogs);
      const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const date = new Intl.DateTimeFormat('en-CA', {
        year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Hong_Kong',
      }).format(new Date());
      const link = document.createElement('a');
      link.href = url;
      link.download = `LPY個案檔案索引_完整備份_${date}.lpybackup`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      const backupAt = backup.createdAt;
      localStorage.setItem(LAST_EXTERNAL_BACKUP_KEY, backupAt);
      if (storageMode === 'indexeddb') await writeMetadata(LAST_EXTERNAL_BACKUP_KEY, backupAt);
      setLastExternalBackupAt(backupAt);
      setToast('完整備份已輸出，請保存到「檔案」或 USB');
    } catch {
      setBackupError('未能建立完整備份，請再試一次。');
    } finally {
      setBackupWorking(false);
    }
  };

  const buildRestoreCandidate = (
    label: string,
    createdAt: string,
    nextSlots: Slot[],
    nextLogs: ActivityLog[],
    source: RestoreCandidate['source'],
  ) => {
    const normalized = normalizeSlots(nextSlots);
    let added = 0;
    let removed = 0;
    let modified = 0;
    normalized.forEach((next, index) => {
      const current = slots[index];
      if (!current.name && next.name) added += 1;
      else if (current.name && !next.name) removed += 1;
      else if (current.name !== next.name
        || current.c !== next.c || current.h !== next.h || current.e !== next.e
        || current.d !== next.d || current.v !== next.v) modified += 1;
    });
    setRestoreCode('');
    setBackupError('');
    setPendingRestore({
      label,
      createdAt,
      slots: normalized,
      activityLogs: [...nextLogs].sort((a, b) => b.timestamp.localeCompare(a.timestamp)).slice(0, 1500),
      source,
      added,
      removed,
      modified,
    });
  };

  const importFullBackup = async (file: File) => {
    setBackupWorking(true);
    setBackupError('');
    try {
      if (file.size > 5 * 1024 * 1024) throw new Error('backup-too-large');
      const parsed: unknown = JSON.parse(await file.text());
      if (!isPortableBackup(parsed) || !(await hasValidChecksum(parsed))) throw new Error('invalid-backup');
      buildRestoreCandidate(
        file.name,
        parsed.createdAt,
        normalizeSlots(parsed.slots),
        parsed.activityLogs,
        'file',
      );
    } catch {
      setBackupError('這個備份檔不完整、已被修改或不是 LPY 完整備份。');
    } finally {
      setBackupWorking(false);
      if (importBackupRef.current) importBackupRef.current.value = '';
    }
  };

  const selectLocalSnapshot = (snapshot: BackupSnapshot) => {
    if (!isValidSlots(snapshot.slots) || !isActivityLogs(snapshot.activityLogs)) {
      setBackupError('這個本機版本資料不完整，不能還原。');
      return;
    }
    buildRestoreCandidate(
      `本機版本 ${formatBackupDate(snapshot.updatedAt)}`,
      snapshot.updatedAt,
      normalizeSlots(snapshot.slots),
      snapshot.activityLogs,
      'snapshot',
    );
  };

  const confirmRestore = async () => {
    if (!pendingRestore || restoreCode !== 'LPY') return;
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    setStorageSaving(false);
    setRestoreWorking(true);
    setBackupError('');
    const restoredAt = new Date().toISOString();
    const nextSlots = pendingRestore.slots.map((slot) => ({ ...slot }));
    const nextLogs = pendingRestore.activityLogs.map((log) => ({ ...log, flags: [...log.flags] }));
    try {
      if (storageMode === 'indexeddb') {
        await createLocalSnapshot({ slots, activityLogs, updatedAt: restoredAt }, 'before-restore');
        await saveCurrentState(
          { slots: nextSlots, activityLogs: nextLogs, updatedAt: restoredAt },
          pendingRestore.source === 'file' ? 'imported' : 'restored',
        );
      }
      localStorage.setItem(STORAGE_KEY, JSON.stringify(nextSlots));
      localStorage.setItem(ACTIVITY_STORAGE_KEY, JSON.stringify(nextLogs));
      localStorage.setItem(LOCAL_UPDATED_AT_KEY, restoredAt);
      lastStoredFingerprint.current = JSON.stringify([nextSlots, nextLogs]);
      setSlots(nextSlots);
      setActivityLogs(nextLogs);
      setLastSavedAt(restoredAt);
      if (storageMode === 'indexeddb') await refreshSnapshots();
      setPendingRestore(null);
      setRestoreCode('');
      setToast('完整資料已安全還原');
    } catch {
      setBackupError('還原失敗，原有資料未有被取代。');
    } finally {
      setRestoreWorking(false);
    }
  };

  return (
    <main className={`app-shell ${view === 'search' ? 'search-mode' : view === 'batch' ? 'batch-mode' : view === 'history' ? 'history-mode' : view === 'backup' ? 'backup-mode' : ''}`}>
      <header className="topbar">
        <div><p className="eyebrow">職業治療</p><h1>個案檔案索引</h1></div>
        <div className="capacity" aria-label={`已使用 ${occupied} 個位置，共 93 個`}><strong>{occupied}</strong><span>/ 93</span></div>
      </header>

      <section className="view-panel" hidden={view !== 'search'}>
        <div className="search-hero neutral-search">
          <label htmlFor="case-search">輸入檔案名稱或標記</label>
          <div className="search-control">
            <div className="search-box" role="combobox" aria-expanded={suggestionsOpen} aria-haspopup="listbox">
              <span className="search-symbol" aria-hidden="true">⌕</span>
              <input
                ref={searchRef}
                id="case-search"
                type="search"
                value={query}
                onChange={(event) => { setQuery(event.target.value); setSuggestionsOpen(true); setSpeechNote(''); }}
                onFocus={() => setSuggestionsOpen(true)}
                onBlur={() => setTimeout(() => setSuggestionsOpen(false), 120)}
                placeholder="姓名或 C／H／E／D／V"
                autoComplete="off"
                enterKeyHint="search"
                aria-autocomplete="list"
                aria-controls="name-suggestions"
              />
              {query && <button className="clear-search" type="button" onClick={() => { setQuery(''); setSuggestionsOpen(true); }} aria-label="清除搜尋">×</button>}
              <button className={`voice-button ${listening ? 'listening' : ''}`} type="button" onClick={startVoiceSearch}
                aria-label={listening ? '停止語音輸入' : '語音輸入'}>●</button>
            </div>
            {suggestionsOpen && suggestions.length > 0 && (
              <div className="suggestions" id="name-suggestions" role="listbox">
                <p>名單選項</p>
                {suggestions.map((name) => (
                  <button key={name} type="button" role="option" onPointerDown={(event) => event.preventDefault()}
                    onClick={() => chooseSuggestion(name)}>{FLAG_ORDER.some((flag) => flag.label === name) ? `${name} 標記` : name}</button>
                ))}
              </div>
            )}
          </div>
          <p>{speechNote || '輸入部分姓名、C／H／E／D／V，或按咪高峰讀出姓名'}</p>
        </div>

        <div className="results" aria-live="polite">
          {!normalizedQuery && (
            <div className="welcome-card neutral-welcome">
              <div className="neutral-index-mark" aria-hidden="true"><i /><i /><i /></div>
              <h2>快速找到個案檔案</h2>
              <p>搜尋後會清楚顯示檔案顏色、編號及 C／H／E／D／V 標記。</p>
              <button type="button" onClick={() => searchRef.current?.focus()}>開始搜尋</button>
            </div>
          )}
          {normalizedQuery && results.length === 0 && (
            <div className="empty-state neutral-welcome">
              <span aria-hidden="true">?</span><h2>找不到「{query.trim()}」</h2>
              <p>請從下拉名單選擇，或到「全部位置」加入新個案。</p>
              <button type="button" onClick={() => setView('files')}>查看空位</button>
            </div>
          )}
          {results.map((slot) => (
            <article className={`result-card ${slot.color}`} key={`${slot.color}-${slot.number}`}>
              <div className="result-label"><span className="color-dot" /><span>{COLOR_INFO[slot.color].label}檔案</span></div>
              <div className="result-main">
                <div><p>個案名稱</p><h2>{slot.name}</h2><FlagBadges slot={slot} /></div>
                <div className="result-number"><span>檔案位置</span><strong>{COLOR_INFO[slot.color].label} {slot.number}</strong></div>
              </div>
            </article>
          ))}
        </div>
      </section>

      <section className="view-panel file-view" hidden={view !== 'files'}>
        <div className="file-heading">
          <div><h2>全部位置</h2><p>每種顏色由 1 至 31 垂直排列。點按空位加入；長按名稱 3 秒管理。</p></div>
          <button type="button" className="export-button" onClick={downloadExcel} disabled={exporting}>
            <span aria-hidden="true">⇩</span>{exporting ? '正在輸出…' : '輸出 Excel 紙本'}
          </button>
        </div>

        <div className="three-file-columns">
          {DISPLAY_COLORS.map((color) => {
            const colorSlots = slots.filter((slot) => slot.color === color);
            const count = colorSlots.filter((slot) => slot.name).length;
            return (
              <section className={`file-section ${color}`} key={color} aria-labelledby={`${color}-heading`}>
                <header className="file-section-heading">
                  <h2 id={`${color}-heading`}>{COLOR_INFO[color].label}</h2>
                  <span>{count}/31</span>
                </header>
                <div className="slot-list">
                  {colorSlots.map((slot) => (
                    <LongPressSlot key={`${slot.color}-${slot.number}`} slot={slot}
                      onAdd={() => openEditor(slot)} onManage={() => openActions(slot)} />
                  ))}
                </div>
              </section>
            );
          })}
        </div>
      </section>

      <section className="view-panel batch-view" hidden={view !== 'batch'}>
        <div className="batch-heading">
          <h2>批量處理</h2>
          <p>一次加入或刪除多位個案。所有更改仍會自動儲存在這部手機。</p>
        </div>
        <div className="batch-tabs" role="tablist" aria-label="批量處理方式">
          <button type="button" role="tab" aria-selected={batchMode === 'add'} className={batchMode === 'add' ? 'active' : ''}
            onClick={() => setBatchMode('add')}>＋ 新增個案</button>
          <button type="button" role="tab" aria-selected={batchMode === 'delete'} className={batchMode === 'delete' ? 'active' : ''}
            onClick={() => setBatchMode('delete')}>－ 刪除個案</button>
        </div>

        {batchMode === 'add' && (
          <form className="batch-card batch-add-card" onSubmit={(event) => { event.preventDefault(); addMultipleCases(); }}>
            <header><div><h3>新增名單</h3><p>尚餘 {emptySlots.length} 個空位</p></div><span className="batch-random-label">預設隨機分配</span></header>
            <div className="batch-add-rows">
              {batchRows.map((row, index) => {
                const selectedElsewhere = new Set(batchRows.filter((item) => item.id !== row.id).map((item) => item.position));
                return (
                  <div className="batch-add-row" key={row.id}>
                    <span className="batch-row-number">{index + 1}</span>
                    <label><span>個案名稱</span><input value={row.name} maxLength={30} placeholder="輸入名稱"
                      onChange={(event) => updateBatchRow(row.id, { name: event.target.value })} /></label>
                    <label><span>檔案位置</span><select value={row.position}
                      onChange={(event) => updateBatchRow(row.id, { position: event.target.value })}>
                      <option value="random">自動隨機分配</option>
                      {emptySlots.filter((slot) => !selectedElsewhere.has(slotKey(slot))).map((slot) => (
                        <option key={slotKey(slot)} value={slotKey(slot)}>{COLOR_INFO[slot.color].label} {slot.number}（空位）</option>
                      ))}
                    </select></label>
                    <fieldset className="batch-tag-picker">
                      <legend>標記（可多選）</legend>
                      {FLAG_ORDER.map((flag) => (
                        <label key={flag.key}><input type="checkbox" checked={row[flag.key]}
                          onChange={(event) => updateBatchRow(row.id, { [flag.key]: event.target.checked })} /><span>{flag.label}</span></label>
                      ))}
                    </fieldset>
                    <button type="button" className="remove-batch-row" onClick={() => removeBatchRow(row.id)} aria-label={`移除第 ${index + 1} 行`}>×</button>
                  </div>
                );
              })}
            </div>
            <button type="button" className="add-row-button" onClick={addBatchRow}>＋ 再加一位</button>
            {batchAddError && <p className="batch-error" role="alert">{batchAddError}</p>}
            <button type="submit" className="batch-primary-button">加入所有已輸入個案</button>
          </form>
        )}

        {batchMode === 'delete' && (
          <div className="batch-card batch-delete-card">
            <header><div><h3>選擇要刪除的個案</h3><p>已選 {batchSelected.length} 位</p></div></header>
            <label className="batch-search"><span aria-hidden="true">⌕</span><input type="search" value={batchDeleteSearch}
              onChange={(event) => setBatchDeleteSearch(event.target.value)} placeholder="搜尋個案名稱" /></label>
            <div className="batch-case-list">
              {filteredDeleteSlots.map((slot) => {
                const key = slotKey(slot);
                return (
                  <label className={`batch-case-option ${slot.color} ${batchSelected.includes(key) ? 'selected' : ''}`} key={key}>
                    <input type="checkbox" checked={batchSelected.includes(key)} onChange={() => toggleBatchSelection(key)} />
                    <span className="batch-checkmark">✓</span>
                    <span className="batch-case-name">{slot.name}<FlagBadges slot={slot} /></span>
                    <strong>{COLOR_INFO[slot.color].label} {slot.number}</strong>
                  </label>
                );
              })}
            </div>
            <button type="button" className="batch-delete-button" disabled={!batchSelected.length}
              onClick={() => { setBatchDeleteCode(''); setBatchDeleteOpen(true); }}>刪除已選 {batchSelected.length} 位個案</button>
          </div>
        )}
      </section>

      <section className="view-panel history-view" hidden={view !== 'history'}>
        <div className="history-heading">
          <h2>新增及刪除紀錄</h2>
          <p>由此功能啟用後開始記錄，日期最新的項目會顯示在最上方。</p>
        </div>

        <section className="history-card add-history" aria-labelledby="add-history-heading">
          <header><div><span aria-hidden="true">＋</span><h3 id="add-history-heading">新增紀錄</h3></div><strong>{addLogs.length} 項</strong></header>
          {addLogs.length ? (
            <div className="history-table" role="table" aria-label="新增個案紀錄">
              <div className="history-table-head" role="row"><span>日期</span><span>個案</span><span>位置／標記</span></div>
              {addLogs.map((log) => (
                <div className="history-row" role="row" key={log.id}>
                  <time dateTime={log.timestamp}>{formatLogDate(log.timestamp)}</time>
                  <strong>{log.name}</strong>
                  <span><em>{COLOR_INFO[log.color].label} {log.number}</em><LogFlagBadges flags={log.flags} /></span>
                </div>
              ))}
            </div>
          ) : <p className="history-empty">暫時未有新增紀錄</p>}
        </section>

        <section className="history-card delete-history" aria-labelledby="delete-history-heading">
          <header><div><span aria-hidden="true">－</span><h3 id="delete-history-heading">刪除紀錄</h3></div><strong>{deleteLogs.length} 項</strong></header>
          {deleteLogs.length ? (
            <div className="history-table" role="table" aria-label="刪除個案紀錄">
              <div className="history-table-head" role="row"><span>日期</span><span>個案</span><span>位置／標記</span></div>
              {deleteLogs.map((log) => (
                <div className="history-row" role="row" key={log.id}>
                  <time dateTime={log.timestamp}>{formatLogDate(log.timestamp)}</time>
                  <strong>{log.name}</strong>
                  <span><em>{COLOR_INFO[log.color].label} {log.number}</em><LogFlagBadges flags={log.flags} /></span>
                </div>
              ))}
            </div>
          ) : <p className="history-empty">暫時未有刪除紀錄</p>}
        </section>
      </section>

      <section className="view-panel backup-view" hidden={view !== 'backup'}>
        <div className="backup-heading">
          <h2>儲存及備份</h2>
          <p>名單會自動保存在這部裝置。請另外下載完整備份，才可跨網站或更換裝置還原。</p>
        </div>

        <section className="storage-status-card" aria-labelledby="storage-status-heading">
          <header>
            <div><span aria-hidden="true">✓</span><h3 id="storage-status-heading">本機儲存狀態</h3></div>
            <strong className={storageMode === 'indexeddb' ? 'ready' : 'limited'}>
              {storageMode === 'loading' ? '檢查中' : storageMode === 'indexeddb' ? 'IndexedDB 已啟用' : '基本儲存'}
            </strong>
          </header>
          <div className="storage-status-grid">
            <div><span>自動儲存</span><strong>{storageSaving ? '儲存中…' : lastSavedAt ? formatBackupDate(lastSavedAt) : '準備中'}</strong></div>
            <div><span>完整備份</span><strong>{lastExternalBackupAt ? formatBackupDate(lastExternalBackupAt) : '尚未建立'}</strong></div>
            <div><span>本機舊版本</span><strong>{snapshots.length} 個</strong></div>
          </div>
        </section>

        <section className="backup-action-card" aria-labelledby="portable-backup-heading">
          <div className="backup-card-copy">
            <span className="backup-card-icon" aria-hidden="true">⇩</span>
            <div><h3 id="portable-backup-heading">完整復原備份</h3><p>JSON 格式，完整保存93個位置、標記及增刪紀錄。</p></div>
          </div>
          <button type="button" className="backup-primary" onClick={downloadFullBackup} disabled={backupWorking}>
            {backupWorking ? '處理中…' : '建立完整備份檔'}
          </button>
          <button type="button" className="backup-secondary" onClick={() => importBackupRef.current?.click()} disabled={backupWorking}>
            匯入及還原完整備份
          </button>
          <input ref={importBackupRef} className="backup-file-input" type="file" accept=".lpybackup,application/json"
            onChange={(event) => { const file = event.target.files?.[0]; if (file) importFullBackup(file); }} />
          <p className="backup-file-note">備份檔副檔名為 <strong>.lpybackup</strong>，可離線保存到「檔案」、USB或電腦。</p>
        </section>

        <section className="backup-action-card excel-backup-card" aria-labelledby="excel-backup-heading">
          <div className="backup-card-copy">
            <span className="backup-card-icon" aria-hidden="true">X</span>
            <div><h3 id="excel-backup-heading">Excel 紙本及查閱</h3><p>保留原有列印版面；完整還原請使用 .lpybackup。</p></div>
          </div>
          <button type="button" className="backup-secondary" onClick={downloadExcel} disabled={exporting}>
            {exporting ? '正在輸出…' : '輸出 Excel 紙本'}
          </button>
        </section>

        <section className="local-version-card" aria-labelledby="local-version-heading">
          <header>
            <div><h3 id="local-version-heading">本機舊版本</h3><p>最近修改自動保留；還原前亦會先保存目前資料。</p></div>
            <button type="button" onClick={saveManualLocalVersion} disabled={backupWorking || storageMode !== 'indexeddb'}>＋ 立即保存</button>
          </header>
          {storageMode !== 'indexeddb' ? (
            <p className="local-version-empty">瀏覽器暫時只可使用基本儲存，請定期建立完整備份檔。</p>
          ) : snapshots.length ? (
            <div className="snapshot-list">
              {snapshots.slice(0, 12).map((snapshot) => (
                <button type="button" key={snapshot.id} onClick={() => selectLocalSnapshot(snapshot)}>
                  <span><strong>{formatBackupDate(snapshot.updatedAt)}</strong><small>{snapshot.source === 'automatic' ? '自動版本'
                    : snapshot.source === 'before-restore' ? '還原前版本'
                      : snapshot.source === 'imported' ? '匯入版本'
                        : snapshot.source === 'restored' ? '已還原版本' : '手動版本'}</small></span>
                  <em>還原</em>
                </button>
              ))}
            </div>
          ) : <p className="local-version-empty">完成下一次修改後，這裏會顯示可還原版本。</p>}
        </section>

        {backupError && <p className="backup-error" role="alert">{backupError}</p>}
        <p className="backup-privacy-note"><strong>私隱提醒：</strong>完整備份及 Excel 均包含個案姓名，請只儲存在獲授權的位置。</p>
      </section>

      <nav className="bottom-nav" aria-label="主要功能">
        <button type="button" className={view === 'search' ? 'active' : ''} onClick={() => setView('search')}><span aria-hidden="true">⌕</span>搜尋</button>
        <button type="button" className={view === 'files' ? 'active' : ''} onClick={() => setView('files')}><span aria-hidden="true">▤</span>全部位置</button>
        <button type="button" className={view === 'batch' ? 'active' : ''} onClick={() => setView('batch')}><span aria-hidden="true">±</span>批量處理</button>
        <button type="button" className={view === 'history' ? 'active' : ''} onClick={() => setView('history')}><span aria-hidden="true">◷</span>紀錄</button>
        <button type="button" className={view === 'backup' ? 'active' : ''} onClick={() => setView('backup')}><span aria-hidden="true">⇩</span>備份</button>
      </nav>

      {selected && (
        <div className="sheet-backdrop" role="presentation" onPointerDown={closeSheet}>
          <section className="action-sheet" role="dialog" aria-modal="true" aria-labelledby="sheet-title"
            onPointerDown={(event) => event.stopPropagation()}>
            <div className="sheet-handle" aria-hidden="true" />
            {sheetMode === 'actions' && (
              <>
                <div className={`sheet-file-summary ${selected.color}`}>
                  <span className="sheet-number">{selected.number}</span>
                  <div><p>{COLOR_INFO[selected.color].label}檔案</p><h2 id="sheet-title">{selected.name}</h2><FlagBadges slot={selected} /></div>
                </div>
                <div className="sheet-actions">
                  <button type="button" className="primary-action" onClick={() => setSheetMode('edit')}>修改名稱及標記</button>
                  <button type="button" className="danger-action" onClick={() => { setDeleteCode(''); setSheetMode('delete'); }}>刪除個案</button>
                  <button type="button" className="cancel-action" onClick={closeSheet}>取消</button>
                </div>
              </>
            )}
            {sheetMode === 'edit' && (
              <form onSubmit={(event) => { event.preventDefault(); saveName(); }}>
                <div className="editor-heading"><p>{COLOR_INFO[selected.color].label}檔案 · {selected.number} 號</p><h2 id="sheet-title">{selected.name ? '修改個案資料' : '加入新個案'}</h2></div>
                <label className="name-field"><span>個案名稱</span><input autoFocus value={draftName}
                  onChange={(event) => { setDraftName(event.target.value); setError(''); }} placeholder="請輸入名稱" maxLength={30} /></label>
                <fieldset className="flag-picker">
                  <legend>標記（可多選）</legend>
                  <label><input type="checkbox" checked={draftC} onChange={(event) => setDraftC(event.target.checked)} /><span>C</span></label>
                  <label><input type="checkbox" checked={draftH} onChange={(event) => setDraftH(event.target.checked)} /><span>H</span></label>
                  <label><input type="checkbox" checked={draftE} onChange={(event) => setDraftE(event.target.checked)} /><span>E</span></label>
                  <label><input type="checkbox" checked={draftD} onChange={(event) => setDraftD(event.target.checked)} /><span>D</span></label>
                  <label><input type="checkbox" checked={draftV} onChange={(event) => setDraftV(event.target.checked)} /><span>V</span></label>
                </fieldset>
                {error && <p className="form-error" role="alert">{error}</p>}
                <div className="editor-actions"><button type="button" onClick={closeSheet}>取消</button><button type="submit">儲存</button></div>
              </form>
            )}
            {sheetMode === 'delete' && (
              <div className="delete-confirm"><span className="warning-mark" aria-hidden="true">!</span>
                <h2 id="sheet-title">刪除「{selected.name}」？</h2>
                <p>刪除後，{COLOR_INFO[selected.color].label} {selected.number} 號會變成空位，所有標記亦會清除。</p>
                <label className="delete-code-field"><span>請輸入確認字樣</span><input value={deleteCode} autoCapitalize="characters" spellCheck={false}
                  onChange={(event) => setDeleteCode(event.target.value.toUpperCase().slice(0, 3))} placeholder="輸入 LPY" /></label>
                <p className="delete-code-hint">確認字樣：<strong>LPY</strong></p>
                <button type="button" className="confirm-delete" disabled={deleteCode !== 'LPY'} onClick={deleteName}>確認刪除</button>
                <button type="button" className="cancel-action" onClick={() => setSheetMode('actions')}>返回</button>
              </div>
            )}
          </section>
        </div>
      )}
      {batchDeleteOpen && (
        <div className="sheet-backdrop batch-delete-backdrop" role="presentation" onPointerDown={() => setBatchDeleteOpen(false)}>
          <section className="action-sheet batch-delete-confirm" role="dialog" aria-modal="true" aria-labelledby="batch-delete-title"
            onPointerDown={(event) => event.stopPropagation()}>
            <div className="sheet-handle" aria-hidden="true" />
            <span className="warning-mark" aria-hidden="true">!</span>
            <h2 id="batch-delete-title">刪除 {batchSelected.length} 位個案？</h2>
            <p>以下個案會被刪除，原有位置將會變成空位：</p>
            <div className="batch-delete-summary">
              {slots.filter((slot) => batchSelected.includes(slotKey(slot))).map((slot) => (
                <div key={slotKey(slot)}><span>{slot.name}</span><strong>{COLOR_INFO[slot.color].label} {slot.number}</strong></div>
              ))}
            </div>
            <label className="delete-code-field"><span>請輸入確認字樣</span><input autoFocus value={batchDeleteCode} autoCapitalize="characters" spellCheck={false}
              onChange={(event) => setBatchDeleteCode(event.target.value.toUpperCase().slice(0, 3))} placeholder="輸入 LPY" /></label>
            <p className="delete-code-hint">確認字樣：<strong>LPY</strong></p>
            <button type="button" className="confirm-delete" disabled={batchDeleteCode !== 'LPY'} onClick={deleteMultipleCases}>確認刪除全部</button>
            <button type="button" className="cancel-action" onClick={() => setBatchDeleteOpen(false)}>取消</button>
          </section>
        </div>
      )}
      {pendingRestore && (
        <div className="sheet-backdrop restore-backdrop" role="presentation" onPointerDown={() => !restoreWorking && setPendingRestore(null)}>
          <section className="action-sheet restore-confirm" role="dialog" aria-modal="true" aria-labelledby="restore-title"
            onPointerDown={(event) => event.stopPropagation()}>
            <div className="sheet-handle" aria-hidden="true" />
            <span className="restore-mark" aria-hidden="true">↶</span>
            <h2 id="restore-title">確認還原完整資料？</h2>
            <p className="restore-source">{pendingRestore.label}<br /><time dateTime={pendingRestore.createdAt}>{formatBackupDate(pendingRestore.createdAt)}</time></p>
            <div className="restore-summary">
              <div><span>備份個案</span><strong>{pendingRestore.slots.filter((slot) => slot.name).length} 位</strong></div>
              <div><span>新增</span><strong>＋{pendingRestore.added}</strong></div>
              <div><span>刪除</span><strong>－{pendingRestore.removed}</strong></div>
              <div><span>修改</span><strong>{pendingRestore.modified}</strong></div>
            </div>
            <p className="restore-safety-note">目前資料會先自動保存為「還原前版本」，完成後仍可返回。</p>
            <label className="delete-code-field"><span>請輸入確認字樣</span><input value={restoreCode} autoCapitalize="characters" spellCheck={false}
              onChange={(event) => setRestoreCode(event.target.value.toUpperCase().slice(0, 3))} placeholder="輸入 LPY" /></label>
            <p className="delete-code-hint">確認字樣：<strong>LPY</strong></p>
            {backupError && <p className="restore-error" role="alert">{backupError}</p>}
            <button type="button" className="confirm-restore" disabled={restoreCode !== 'LPY' || restoreWorking} onClick={confirmRestore}>
              {restoreWorking ? '正在還原…' : '確認完整還原'}
            </button>
            <button type="button" className="cancel-action" disabled={restoreWorking} onClick={() => { setPendingRestore(null); setRestoreCode(''); setBackupError(''); }}>取消</button>
          </section>
        </div>
      )}
      {toast && <div className="toast" role="status">✓ {toast}</div>}
    </main>
  );
}
