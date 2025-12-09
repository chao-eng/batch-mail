import { contextBridge, ipcRenderer } from "electron";

/*
暴露homepageWindow窗口主进程的方法到homepageWindow窗口的渲染进程
*/
contextBridge.exposeInMainWorld("homepageWindowAPI", {
  setConfig: (key, value) => ipcRenderer.invoke("setConfig", key, value),
  getConfig: (key) => ipcRenderer.invoke("getConfig", key),
  checkConfig: (key) => ipcRenderer.invoke("checkConfig", key),
  sendMail: (mailData) => ipcRenderer.invoke("sendMail", mailData),
  parseExcel: (fileData) => ipcRenderer.invoke('parseExcel', fileData),
  startBatchTasks: (tasks) => ipcRenderer.invoke('startBatchTasks', tasks),
  onBatchUpdate: (callback) => ipcRenderer.on('batch-update', callback),
  removeBatchUpdateListener: () => ipcRenderer.removeAllListeners('batch-update'),
  downloadTemplate: () => ipcRenderer.invoke('downloadTemplate'),
});
