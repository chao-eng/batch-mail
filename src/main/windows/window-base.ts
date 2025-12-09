import { app, BrowserWindow, IpcMainEvent, IpcMainInvokeEvent, BrowserWindowConstructorOptions, ipcMain } from "electron";

/**
 * 窗口基类，所有的窗口都继承自该类，如 PrimaryWindow、FramelessWindow
 * @class
 */
abstract class WindowBase {

  protected _browserWindow: BrowserWindow | null = null;
  protected _ipcHandlers: string[] = []; // 记录注册的 IPC 事件名
  protected _ipcHandleHandlers: string[] = []; // 记录 ipcMain.handle 事件
  constructor(options?: BrowserWindowConstructorOptions) {
    this._browserWindow = new BrowserWindow(options);

    if (this._browserWindow) {
      // After received closed event, remove the reference to the window and avoid using it any more.
      this._browserWindow.on("closed", () => {
        this.cleanupIpcHandlers();
        this._browserWindow = null;
      });
    }

    this.registerIpcMainHandler();
  }
  // 新增：清理 IPC 处理器的方法
  protected cleanupIpcHandlers(): void {
    // 清理 ipcMain.on 注册的事件
    this._ipcHandlers.forEach(eventName => {
      console.log(`[Electron] Removing IPC handler: ${eventName}`);
      ipcMain.removeAllListeners(eventName);
    });

    // 清理 ipcMain.handle 注册的事件（关键！）
    this._ipcHandleHandlers.forEach(eventName => {
      console.log(`[Electron] Removing IPC handle handler: ${eventName}`);
      ipcMain.removeHandler(eventName); // 使用 removeHandler 而不是 removeAllListeners
    });

    this._ipcHandlers = [];
    this._ipcHandleHandlers = [];
  }
  public openRouter(routerPath: string) {
    let url = "";
    if (app.isPackaged) {
      url = `file://${app.getAppPath()}/build/renderer/index.html#${routerPath}`;
    } else {
      const rendererPort = process.argv[2];
      url = `http://localhost:${rendererPort}/#${routerPath}`;
    }

    console.log(`Load URL: ${url}`);

    if (this._browserWindow) {
      this._browserWindow.loadURL(url);
    }
  }

  public get valid() {
    return this.browserWindow != null;
  }

  public get browserWindow() {
    return this._browserWindow;
  }

  // 新增：注册 IPC 处理器的辅助方法
  protected registerIpcHandler(eventName: string, handler: (...args: any[]) => void): void {
    // 检查是否已经注册过，避免重复
    if (!this._ipcHandlers.includes(eventName)) {
      ipcMain.on(eventName, handler);
      this._ipcHandlers.push(eventName);
    }
  }

  protected registerIpcHandleHandler(eventName: string, handler: (...args: any[]) => any): void {
    // 检查是否已经注册过，避免重复
    if (!this._ipcHandleHandlers.includes(eventName)) {
      ipcMain.handle(eventName, handler);
      this._ipcHandleHandlers.push(eventName);
    }
  }

  protected abstract registerIpcMainHandler(): void;



  public isIpcMainEventBelongMe(event: IpcMainEvent | IpcMainInvokeEvent): boolean {
    if (!this._browserWindow)
      return false;
    return (event.sender.id == this.browserWindow?.webContents.id);
  }
}

export default WindowBase;