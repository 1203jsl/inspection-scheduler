import { app, BrowserWindow, dialog, ipcMain, shell } from 'electron';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { CurrentUser, DataFiles, Result, Settings } from '../shared/types';
import { DataStore } from './core/dataStore';
import { buildTemplate } from './core/excel/template';
import { ImportService } from './core/importService';
import { templateDef } from './core/schema';
import { AppError, toErrorInfo } from './core/errors';
import { LockManager } from './core/lock';
import { readLocalConfig, writeLocalConfig } from './core/localConfig';

// 날짜 입력 칸 등을 한국어 형식으로 (영문 Windows 에서도 동일하게)
app.commandLine.appendSwitch('lang', 'ko-KR');

// ───────── 상태 ─────────
const account = os.userInfo().username;
const host = os.hostname();
let displayName = account;
let dataDir: string | null = null;
let store: DataStore | null = null;
let lock: LockManager | null = null;
let importer: ImportService | null = null;
let lockStaleMinutes = 30;
let win: BrowserWindow | null = null;

const localDir = () => app.getPath('userData');
const me = (): CurrentUser => ({ account, displayName, host });

function openStore(dir: string) {
  dataDir = dir;
  store = new DataStore(dir, { account, displayName });
  lock = new LockManager(dir, me(), () => lockStaleMinutes);
  importer = new ImportService(store, () => ({ account, displayName }));
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

// ───────── 엑셀 ─────────
handle('template:save', async () => {
  const s = await need(store).load('settings');
  const def = templateDef(s);
  const r = await dialog.showSaveDialog(win!, {
    title: '빈 양식 저장',
    defaultPath: path.join(app.getPath('downloads'), `검수신청_양식_${def.version}.xlsx`),
    filters: [{ name: 'Excel 통합 문서', extensions: ['xlsx'] }],
  });
  if (r.canceled || !r.filePath) return null;
  try {
    await fs.writeFile(r.filePath, await buildTemplate(def, s.inspectionItems));
  } catch (e) {
    const code = (e as NodeJS.ErrnoException).code;
    if (code === 'EBUSY' || code === 'EPERM') {
      throw new AppError('FILE_BUSY', `같은 이름의 파일이 엑셀에서 열려 있어 저장하지 못했습니다. 파일을 닫거나 다른 이름으로 저장하세요.`);
    }
    throw e;
  }
  return r.filePath;
});

handle('import:pick', async () => {
  const r = await dialog.showOpenDialog(win!, {
    title: '업체가 보낸 엑셀 선택',
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Excel 통합 문서', extensions: ['xlsx'] }],
  });
  return r.canceled ? [] : r.filePaths;
});
handle('import:preview', async (paths: string[]) => need(importer).preview(paths));
handle('import:commit', async (keys: string[]) => need(importer).commit(keys));

/** 데이터 폴더 안의 파일 열기 (원본 엑셀 등) */
handle('file:open', async (rel: string) => {
  const root = need(dataDir);
  const full = path.resolve(root, rel);
  if (!full.startsWith(path.resolve(root) + path.sep)) throw new AppError('UNKNOWN', '데이터 폴더 밖의 파일은 열 수 없습니다.');
  const err = await shell.openPath(full);
  if (err) throw new AppError('UNKNOWN', `파일을 열 수 없습니다: ${err}`);
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
