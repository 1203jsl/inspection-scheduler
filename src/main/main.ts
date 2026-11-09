import { app, BrowserWindow, dialog, ipcMain } from 'electron';
import os from 'node:os';
import path from 'node:path';
import type { CurrentUser, DataFiles, Result, Settings } from '../shared/types';
import { DataStore } from './core/dataStore';
import { AppError, toErrorInfo } from './core/errors';
import { LockManager } from './core/lock';
import { readLocalConfig, writeLocalConfig } from './core/localConfig';

// ───────── 상태 ─────────
const account = os.userInfo().username;
const host = os.hostname();
let displayName = account;
let dataDir: string | null = null;
let store: DataStore | null = null;
let lock: LockManager | null = null;
let lockStaleMinutes = 30;
let win: BrowserWindow | null = null;

const localDir = () => app.getPath('userData');
const me = (): CurrentUser => ({ account, displayName, host });

function openStore(dir: string) {
  dataDir = dir;
  store = new DataStore(dir, { account, displayName });
  lock = new LockManager(dir, me(), () => lockStaleMinutes);
}

function applySettings(s: Settings) {
  displayName = s.userMap?.[account] || account;
  lockStaleMinutes = Number(s.lockStaleMinutes) || 30;
  store?.setUser({ account, displayName });
  if (dataDir) lock = new LockManager(dataDir, me(), () => lockStaleMinutes);
}

function need<T>(v: T | null, what = '데이터 폴더'): T {
  if (!v) throw new AppError('NOT_INITIALIZED', `${what}가 선택되지 않았습니다. 데이터 폴더를 먼저 선택하세요.`);
  return v;
}

/** IPC 처리: 예외를 화면용 Result 로 바꿉니다. */
function handle<A extends unknown[], T>(ch: string, fn: (...a: A) => Promise<T>) {
  ipcMain.handle(ch, async (_e, ...a: A): Promise<Result<T>> => {
    try {
      return { ok: true, value: await fn(...a) };
    } catch (e) {
      return { ok: false, error: toErrorInfo(e) };
    }
  });
}

// ───────── IPC ─────────
handle('app:getState', async () => ({ dataDir, user: me() }));

handle('folder:pick', async () => {
  const r = await dialog.showOpenDialog(win!, {
    title: '데이터 폴더 선택',
    properties: ['openDirectory', 'createDirectory'],
  });
  return r.canceled || !r.filePaths[0] ? null : r.filePaths[0];
});

handle('folder:check', async (dir: string) => DataStore.checkFolder(dir));

handle('folder:setup', async (opts: { path: string; mode: 'existing' | 'new'; displayName: string }) => {
  const chk = await DataStore.checkFolder(opts.path);
  if (!chk.exists) throw new AppError('FOLDER_MISSING', `폴더를 찾을 수 없습니다 (${opts.path}). 폴더를 다시 선택하세요.`);
  if (opts.mode === 'existing' && !chk.hasData) {
    throw new AppError('NOT_INITIALIZED', '이 폴더에는 검수일정 데이터가 없습니다. 폴더를 다시 고르거나 "새 데이터 폴더 만들기"를 선택하세요.');
  }
  if (opts.mode === 'new' && chk.hasData) {
    throw new AppError('UNKNOWN', '이 폴더에는 이미 검수일정 데이터가 있습니다. "이미 쓰고 있는 데이터 폴더에 연결"을 선택하세요.');
  }
  const name = opts.displayName.trim() || account;
  displayName = name;
  if (opts.mode === 'new') await DataStore.initFolder(opts.path, { account, displayName: name });
  openStore(opts.path);
  const s = await store!.load('settings');
  if (s.userMap?.[account] !== name) {
    const saved = await store!.save({ settings: { ...s, userMap: { ...s.userMap, [account]: name } } });
    applySettings(saved.settings!);
  } else applySettings(s);
  await writeLocalConfig(localDir(), { dataDir: opts.path });
  return { dataDir, user: me() };
});

handle('data:loadAll', async () => {
  const all = await need(store).loadAll();
  applySettings(all.settings);
  return all;
});

handle('data:save', async (changes: Partial<DataFiles>) => {
  const saved = await need(store).save(changes);
  if (saved.settings) applySettings(saved.settings);
  return saved;
});

handle('lock:status', async () => need(lock).read());
handle('lock:acquire', async () => need(lock).acquire());
handle('lock:release', async () => need(lock).release());

handle('backup:list', async () => need(store).listBackups());
handle('backup:restore', async (id: string) => {
  await need(store).restoreBackup(id);
  return true;
});

// ───────── 창 ─────────
async function createWindow() {
  const cfg = await readLocalConfig(localDir());
  if (cfg.dataDir) {
    openStore(cfg.dataDir);
    try {
      applySettings(await store!.load('settings'));
    } catch {
      /* 폴더가 끊겼어도 화면에서 안내합니다 */
    }
  }
  win = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1280,
    minHeight: 720,
    title: '검수일정 관리',
    backgroundColor: '#f3f5f9',
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  await win.loadFile(path.join(__dirname, '../renderer/index.html'));
}

// 편집 중 표시 갱신 (1분마다)
setInterval(() => {
  lock?.heartbeat().catch(() => {});
}, 60_000);

app.whenReady().then(createWindow);
app.on('window-all-closed', () => app.quit());
let releasing = false;
app.on('before-quit', (e) => {
  if (releasing || !lock) return;
  e.preventDefault();
  releasing = true;
  lock.release().catch(() => {}).finally(() => app.quit());
});
