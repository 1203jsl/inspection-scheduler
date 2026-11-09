import { contextBridge, ipcRenderer } from 'electron';

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
};

contextBridge.exposeInMainWorld('api', api);
