import { contextBridge, ipcRenderer, webUtils } from 'electron';

const call = (ch: string) => (...args: unknown[]) => ipcRenderer.invoke(ch, ...args);

/** 화면(renderer)에서 window.api 로 사용 */
const api = {
  getState: call('app:getState'),
  pickFolder: call('folder:pick'),
  checkFolder: call('folder:check'),
  setupFolder: call('folder:setup'),
  loadAll: call('data:loadAll'),
  save: call('data:save'),
  lockStatus: call('lock:status'),
  lockAcquire: call('lock:acquire'),
  lockRelease: call('lock:release'),
  listBackups: call('backup:list'),
  restoreBackup: call('backup:restore'),
  saveTemplate: call('template:save'),
  pickImportFiles: call('import:pick'),
  importPreview: call('import:preview'),
  importCommit: call('import:commit'),
  openDataFile: call('file:open'),
  /** 끌어다 놓은 파일의 실제 경로 */
  pathForFile: (f: File) => webUtils.getPathForFile(f),
};

contextBridge.exposeInMainWorld('api', api);
