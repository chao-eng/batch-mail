/**
 * @file 与系统托盘的相关的功能
 */
import { app } from "electron";
import path from "path";
import { Menu, MenuItem, Tray, nativeImage } from "electron";
import appState, { AppEnv } from "./app-state";

// 创建系统托盘
function CreateAppTray(): Tray {
  const iconPath = process.platform === "win32" ?
    path.join(appState.mainStaticPath, "tray.ico") :
    path.join(appState.mainStaticPath, "tray.png");

  let icon = nativeImage.createFromPath(iconPath);
  // macOS 托盘图标通常需要较小的尺寸 (如 16x16 或 22x22)
  if (process.platform === 'darwin') {
    icon = icon.resize({ width: 20, height: 20 });
    icon.setTemplateImage(true); // 设置为模板图像以适应暗黑模式
  }

  const tray = new Tray(icon);

  tray.on("click", () => {
    appState.homepageWindow?.browserWindow?.show();
  });

  // 创建托盘右键菜单
  const contextMenu = Menu.buildFromTemplate([
    {
      label: "打开",
      type: "normal",
      accelerator: "Alt+O",
      registerAccelerator: true,
      click: () => {
        const win = appState.homepageWindow?.browserWindow;
        if (win) {
          if (win.isVisible()) {
            if (win.isMinimized()) {
              win.restore();
            }
          }
          else {
            win.show();
          }
        }
      },
    },
    {
      label: "退出",
      type: "normal",
      click: () => {
        app.quit();
      },
    },
  ]);
  // 在非生产环境添加一个打开调试工具菜单，方便调试
  if (appState.appEnv != AppEnv.Production) {
    contextMenu.insert(
      0,
      new MenuItem({
        label: "打开开发者工具",
        type: "normal",
        accelerator: "Alt+D",
        registerAccelerator: true,
        click: () => {
          appState.homepageWindow?.browserWindow?.webContents.openDevTools();
        },
      }),
    );
  }

  tray.setContextMenu(contextMenu);
  tray.setToolTip("BatchMail");
  // tray.setTitle("BM");

  return tray;
}

export { CreateAppTray };
