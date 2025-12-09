import { BrowserWindow, app, dialog, session, Menu, net } from "electron";
import log from "electron-log/main";
import PrimaryWindow from "./windows/primary";
import HomepageWindow from "./windows/homepage";
import { CreateAppTray } from "./tray";
import appState from "./app-state";
import { protocol } from "electron";
import fs from "fs";
import path from "path";
import { pathToFileURL } from "url";

// 在 app 启动的最早期执行
app.setName('BatchMail');


// 禁用沙盒
// 在某些系统环境上，不禁用沙盒会导致界面花屏
// app.commandLine.appendSwitch("no-sandbox");
//注册 协议
protocol.registerSchemesAsPrivileged([
  {
    scheme: 'local-file',
    privileges: {
      standard: true,
      secure: true,
      bypassCSP: true,
      corsEnabled: true,
      supportFetchAPI: true
    }
  }
]);
// 创建基本菜单模板，包含复制粘贴等功能
const createMenuTemplate = () => {
  const template: Electron.MenuItemConstructorOptions[] = [];

  if (process.platform === 'darwin') {
    // macOS 菜单
    template.push({
      label: app.getName(),
      submenu: [
        { role: 'about' },
        { type: 'separator' },
        { role: 'services' },
        { type: 'separator' },
        { role: 'hide' },
        { role: 'hideOthers' },
        { role: 'unhide' },
        { type: 'separator' },
        { role: 'quit' }
      ]
    });
  }

  // 编辑菜单 - 这是关键部分
  template.push({
    label: '编辑',
    submenu: [
      { role: 'undo', label: '撤销' },
      { role: 'redo', label: '重做' },
      { type: 'separator' },
      { role: 'cut', label: '剪切' },
      { role: 'copy', label: '复制' },
      { role: 'paste', label: '粘贴' },
      { role: 'pasteAndMatchStyle', label: '粘贴并匹配样式' },
      { role: 'delete', label: '删除' },
      { role: 'selectAll', label: '全选' }
    ]
  });

  // 窗口菜单
  template.push({
    label: '窗口',
    submenu: [
      { role: 'minimize', label: '最小化' },
      { role: 'close', label: '关闭' }
    ]
  });

  return template;
};

// 设置菜单
const menu = Menu.buildFromTemplate(createMenuTemplate());
// 移除默认菜单栏
Menu.setApplicationMenu(menu);

function setupContextMenu(window: BrowserWindow) {
  window.webContents.on('context-menu', (e, params) => {
    const menuTemplate = [
      {
        label: '复制',
        click: () => window.webContents.copy(),
        visible: params.editFlags.canCopy
      },
      {
        label: '粘贴',
        click: () => window.webContents.paste(),
        visible: params.editFlags.canPaste
      }
    ];

    const contextMenu = Menu.buildFromTemplate(menuTemplate);
    contextMenu.popup({ window });
  });
}

const gotLock = app.requestSingleInstanceLock();

// 如果程序只允许启动一个实例时，第二个实例启动后会直接退出
if (!gotLock && appState.onlyAllowSingleInstance) {
  app.quit();
} else {
  app.whenReady().then(async () => {
    if (!appState.initialize()) {
      dialog.showErrorBox("App initialization failed", "The program will exit after click the OK button.",);
      app.exit();
      return;
    }

    log.info("App initialize ok");

    // 注册本地文件协议处理器
    protocol.handle('local-file', async (request) => {
      try {
        console.log('[Electron] 处理本地文件请求:', request.url);

        // 移除协议前缀
        let filePath = request.url.slice('local-file://'.length);

        // URL 解码
        filePath = decodeURIComponent(filePath);

        // 平台特定的路径修正
        if (process.platform === 'win32') {
          // Windows: 确保有盘符
          if (!filePath.match(/^[A-Za-z]:/)) {
            // 如果没有盘符，可能需要添加，但这里先报错
            console.error('[Electron] Windows 路径缺少盘符:', filePath);
            return new Response('Invalid Windows path', { status: 400 });
          }
        } else {
          // Unix/macOS: 确保以 / 开头
          if (!filePath.startsWith('/')) {
            filePath = '/' + filePath;
          }
        }

        console.log('[Electron] 最终文件路径:', filePath);

        // 检查文件是否存在
        if (!fs.existsSync(filePath)) {
          console.error('[Electron] 文件不存在:', filePath);
          return new Response('File not found', { status: 404 });
        }

        // 读取文件
        const fileBuffer = fs.readFileSync(filePath);

        // 根据文件扩展名设置 MIME 类型
        const ext = path.extname(filePath).toLowerCase();
        const mimeTypes: { [key: string]: string } = {
          '.md': 'text/markdown',
          '.txt': 'text/plain',
          '.html': 'text/html',
          '.css': 'text/css',
          '.js': 'application/javascript',
          '.json': 'application/json',
          '.png': 'image/png',
          '.jpg': 'image/jpeg',
          '.jpeg': 'image/jpeg',
          '.gif': 'image/gif',
          '.svg': 'image/svg+xml',
        };

        const contentType = mimeTypes[ext] || 'application/octet-stream';

        return new Response(fileBuffer, {
          headers: {
            'Content-Type': contentType,
            'Cache-Control': 'no-cache'
          }
        });

      } catch (error) {
        console.error('[Electron] 处理本地文件协议时出错:', error);
        return new Response('Internal error', { status: 500 });
      }
    });

    //启动窗口
    //appState.primaryWindow = new PrimaryWindow();
    appState.homepageWindow = new HomepageWindow();
    setupContextMenu(<BrowserWindow>appState.homepageWindow.browserWindow);
    appState.tray = CreateAppTray();

    session.defaultSession.webRequest.onHeadersReceived((details, callback) => {
      callback({
        responseHeaders: {
          ...details.responseHeaders,
          "Content-Security-Policy":
            ["script-src 'self'  'unsafe-inline' https://unpkg.com  img-src  local-file: data: blob: media-src  local-file:;"],
        },
      });
    });
  });

  // 当程序的第二个实例启动时，显示第一个实例的主窗口
  app.on("second-instance", () => {
    appState.homepageWindow?.browserWindow?.show();
  });

  app.on("activate", () => {
    // 在 macOS 中，当
    // 当单击 dock 图标且没有打开其他窗口时，通常会重新创建一个窗口。
    if (BrowserWindow.getAllWindows().length === 0)
      appState.homepageWindow = new HomepageWindow();
  });

  app.on("window-all-closed", () => {
    if (process.platform !== "darwin")
      app.quit();
  });

  app.on("will-quit", () => {
    appState.uninitialize();
  });
}