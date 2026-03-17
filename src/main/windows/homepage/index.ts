import path from "path";
import WindowBase from "../window-base";
import configManager from "./config-manager";
import * as XLSX from 'xlsx';
import nodemailer, { Transporter } from 'nodemailer';
import * as fs from 'fs';
import { dialog, app } from 'electron';
import puppeteer from 'puppeteer-core';
import log from "electron-log/main";
import appState from "../../app-state";

// 定义配置接口
interface SmtpConfig {
  host: string;
  port: number;
  secure: boolean;
  auth: {
    user: string;
    pass: string;
  };
}

// 延时函数
const delay = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

class homepageWindow extends WindowBase {
  private activeBrowser: any = null;

  constructor() {
    const iconPath = process.platform === "win32" ?
      path.join(appState.mainStaticPath, "tray.ico") :
      path.join(appState.mainStaticPath, "tray.png");

    super({
      width: 1000,
      height: 600,
      icon: iconPath,
      webPreferences: {
        preload: path.join(__dirname, "preload.js"),
      },
    });

    this.openRouter("/homepage");
  }

  protected registerIpcMainHandler(): void {
    // 获取配置文件
    this.registerIpcHandleHandler('getConfig', (event, key) => configManager.get(key))

    // 保存配置文件
    this.registerIpcHandleHandler('setConfig', async (event, key, value) => {
      await configManager.set(key, value);
      log.info("setConfig", key, value);
      log.info("保存目录: ", await configManager.getPath());
    });

    // 测试连接
    this.registerIpcHandleHandler('checkConfig', async (event, rawConfig) => {
      log.info("checkConfig", rawConfig);
      try {
        const config: SmtpConfig = {
          host: rawConfig.smtp_server,
          port: rawConfig.smtp_port,
          secure: true,
          auth: {
            user: rawConfig.sender_email,
            pass: rawConfig.password,
          },
        };
        log.info('🚀 开始测试 SMTP 连接 (TypeScript Mode)...');
        const transporter: Transporter = nodemailer.createTransport(config);
        const verifySuccess: boolean = await transporter.verify();

        if (verifySuccess) {
          log.info('✅ [连接成功] 服务器配置正确，身份验证通过！');
          return { status: true, msg: "服务器配置正确，身份验证通过" }
        }
      } catch (error: any) {
        log.error('❌ [发生错误]');
        if (error instanceof Error) {
          log.error(`   错误信息: ${error.message}`);
          if ((error as any).code === 'EAUTH') {
            log.warn('   👉 提示: 认证失败。请检查邮箱账号或应用专用密码。');
            return { status: false, msg: "👉 提示: 认证失败。请检查邮箱账号或应用专用密码。" }
          }
        } else {
          log.error('   未知错误:', error);
        }
        return { status: false, msg: error.message }
      }
    })

    // 发送邮件
    this.registerIpcHandleHandler('sendMail', async (event, params) => {
      log.info('发送邮件:', params)
      const { smtp_server, smtp_port, sender_email, password, to, cc, subject, text, html, attachments } = params;
      const transporter = nodemailer.createTransport({
        host: smtp_server,
        port: parseInt(smtp_port),
        secure: parseInt(smtp_port) === 465,
        auth: {
          user: sender_email,
          pass: password,
        },
      });

      try {
        const info = await transporter.sendMail({
          from: sender_email,
          to: to,
          cc: cc,
          subject: subject,
          text: text,
          attachments: attachments,
          html: html || text
        });
        return { status: true, msg: '发送成功', id: info.messageId };
      } catch (error: any) {
        return { status: false, msg: error.message };
      }
    })

    // 解析 Excel
    this.registerIpcHandleHandler('parseExcel', async (event, payload) => {
      const { fileData } = payload;
      log.info('📄 正在解析 Excel...');
      try {
        const workbook = XLSX.read(fileData, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const rows: any[][] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
        const tasks: any[] = [];
        rows.forEach((row, index) => {
          if (index === 0) return;
          const receiver = row[0];
          const username = row[1] || '';
          if (receiver && typeof receiver === 'string' && receiver.includes('@')) {
            tasks.push({
              id: `task-${Date.now()}-${index}`,
              receiver: receiver,
              username: username,
              status: 'pending',
              error: ''
            });
          }
        });
        return { status: true, data: tasks, msg: `解析成功，共 ${tasks.length} 个任务` };
      } catch (error: any) {
        return { status: false, msg: '解析失败: ' + error.message };
      }
    });

    // 执行批量任务 (SMTP)
    this.registerIpcHandleHandler('startBatchTasks', async (event, tasks: any[]) => {
      log.info(`🚀 开始执行 ${tasks.length} 个任务`);
      const config = await configManager.get('config');
      if (!config?.smtp_server) {
        return { status: false, msg: 'SMTP 配置未找到' };
      }
      const template = await configManager.get('email_template');
      if (!template || !template.subject) {
        return { status: false, msg: '邮件模板未配置（请先设置主题和内容）' };
      }
      const transporter = nodemailer.createTransport({
        host: config.smtp_server,
        port: parseInt(config.smtp_port),
        secure: parseInt(config.smtp_port) === 465,
        auth: { user: config.sender_email, pass: config.password },
      });
      const templateWithConfig = { ...template, _config: config };
      this.processQueue(event.sender, transporter, tasks, config.sender_email, templateWithConfig, config.cc_emails);
      return { status: true, msg: '任务队列已启动' };
    });

    // 浏览器发送任务
    this.registerIpcHandleHandler('startBrowserTasks', async (event, tasks: any[]) => {
      log.info(`🚀 开始执行浏览器发送任务: ${tasks.length} 个任务`);
      const config = await configManager.get('config');
      const template = await configManager.get('email_template');
      if (!template || !template.subject) {
        return { status: false, msg: '邮件模板未配置（请先设置主题和内容）' };
      }
      const templateWithConfig = { ...template, _config: config };
      this.runBrowserBatch(event.sender, tasks, templateWithConfig);
      return { status: true, msg: '浏览器发送任务已启动' };
    });

    // 下载模板
    this.registerIpcHandleHandler('downloadTemplate', async () => {
      const headers = [['收件人', '用户名']];
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet(headers);
      worksheet['!cols'] = [{ wch: 30 }, { wch: 20 }];
      XLSX.utils.book_append_sheet(workbook, worksheet, '批量发送模板');
      const buffer = XLSX.write(workbook, { type: 'buffer' });
      const { canceled, filePath } = await dialog.showSaveDialog({
        title: '保存 Excel 模板',
        defaultPath: '邮件批量发送模板.xlsx',
        filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
      });
      if (canceled || !filePath) return { status: false, msg: '取消下载' };
      try {
        fs.writeFileSync(filePath, buffer);
        return { status: true, msg: '模板已保存' };
      } catch (error: any) {
        return { status: false, msg: '保存失败: ' + error.message };
      }
    });

    // 导出发送结果
    this.registerIpcHandleHandler('exportResults', async (event, data: any[]) => {
      const exportData = data.map(item => ({
        '收件人': item.receiver,
        '用户名': item.username,
        '状态': item.status === 'success' ? '成功' : (item.status === 'failed' ? '失败' : '待发送'),
        '失败原因': item.error || ''
      }));
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.json_to_sheet(exportData);
      worksheet['!cols'] = [{ wch: 30 }, { wch: 20 }, { wch: 15 }, { wch: 40 }];
      XLSX.utils.book_append_sheet(workbook, worksheet, '发送结果');
      const buffer = XLSX.write(workbook, { type: 'buffer' });
      const { canceled, filePath } = await dialog.showSaveDialog({
        title: '导出发送结果',
        defaultPath: `发送结果_${new Date().toISOString().split('T')[0]}.xlsx`,
        filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
      });
      if (canceled || !filePath) return { status: false, msg: '取消导出' };
      try {
        fs.writeFileSync(filePath, buffer);
        return { status: true, msg: '结果已导出' };
      } catch (error: any) {
        return { status: false, msg: '导出保存失败: ' + error.message };
      }
    });

    // 账户登录
    this.registerIpcHandleHandler('openLoginBrowser', async (event) => {
      try {
        const browser = await this.getBrowser();
        const pages = await browser.pages();
        const page = pages.length > 0 ? pages[0] : await browser.newPage();
        log.info('[Puppeteer Login] 正在打开 Gmail...');
        await page.goto('https://mail.google.com/mail/?authuser=0', { waitUntil: 'networkidle2', timeout: 0 });

        return new Promise((resolve) => {
          if (page.url().includes('mail.google.com')) {
            delay(2000).then(async () => {
              // 注意：为了复用，这里不再关闭浏览器，只置顶窗口
              try {
                const session = await page.target().createCDPSession();
                await session.send('Page.bringToFront');
              } catch (e) {}
              resolve({ status: true, msg: '账户已在登录状态' });
            });
            return;
          }
          const intervalId = setInterval(async () => {
            try {
              const url = page.url();
              if (url.includes('mail.google.com')) {
                clearInterval(intervalId);
                log.info('[Puppeteer Login] 检测到进入 Gmail 域名，登录成功');
                await delay(2000);
                resolve({ status: true, msg: '登录成功' });
              }
            } catch (e) {
              clearInterval(intervalId);
              resolve({ status: false, msg: '浏览器已关闭' });
            }
          }, 1000);
          
          const onDisconnected = () => {
            clearInterval(intervalId);
            resolve({ status: false, msg: '浏览器已关闭' });
          };
          browser.once('disconnected', onDisconnected);
        });
      } catch (err: any) {
        log.error('[Puppeteer Login] 启动错误:', err.message);
        return { status: false, msg: err.message };
      }
    });
  }

  // ------------------------------------------------
  // 核心方法：获取或创建浏览器实例 (实现复用)
  // ------------------------------------------------
  private async getBrowser() {
    if (this.activeBrowser) {
      try {
        // 尝试获取页面以检查连接是否依然有效
        await this.activeBrowser.pages();
        return this.activeBrowser;
      } catch (e) {
        log.warn('[Puppeteer] 原有浏览器实例已失效，准备重启...');
        this.activeBrowser = null;
      }
    }

    let chromePath = "";
    if (process.platform === 'darwin') {
      chromePath = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
    } else if (process.platform === 'win32') {
      const potentialPaths = [
        'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
        path.join(process.env.LOCALAPPDATA || "", "Google/Chrome/Application/chrome.exe")
      ];
      chromePath = potentialPaths.find(p => fs.existsSync(p)) || "";
    }
    if (!chromePath || !fs.existsSync(chromePath)) {
      throw new Error('未发现 Google Chrome 浏览器，请先安装 Chrome。');
    }
    const userDataPath = path.join(app.getPath('userData'), 'chrome-profile');
    if (!fs.existsSync(userDataPath)) {
      fs.mkdirSync(userDataPath, { recursive: true });
    }

    log.info('[Puppeteer] 正在启动系统 Chrome...', chromePath);
    this.activeBrowser = await puppeteer.launch({
      executablePath: chromePath,
      userDataDir: userDataPath,
      headless: false,
      defaultViewport: null,
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized'],
      ignoreDefaultArgs: ['--enable-automation']
    });

    this.activeBrowser.on('disconnected', () => {
      log.info('[Puppeteer] 浏览器已断开连接');
      this.activeBrowser = null;
    });

    return this.activeBrowser;
  }


  // ------------------------------------------------
  // 辅助方法：处理队列 (SMTP)
  // ------------------------------------------------
  async processQueue(sender: Electron.WebContents, transporter: any, tasks: any[], fromEmail: string, template: any, ccEmails?: string) {
    for (const task of tasks) {
      if (task.status === 'success') continue;
      sender.send('batch-update', { id: task.id, status: 'processing' });
      try {
        const replaceVars = (str: string) => {
          if (!str) return '';
          return str.replace(/\{\{username\}\}/g, task.username || '');
        };
        const subject = replaceVars(template.subject);
        const html = replaceVars(template.htmlContent || template.content);
        await transporter.sendMail({
          from: fromEmail,
          to: task.receiver,
          cc: ccEmails,
          subject: subject,
          html: html
        });
        sender.send('batch-update', { id: task.id, status: 'success' });
        log.info(`✅ [${task.receiver}] 发送成功`);
      } catch (err: any) {
        sender.send('batch-update', { id: task.id, status: 'failed', error: err.message });
        log.error(`❌ [${task.receiver}] 发送失败: ${err.message}`);
      }
      const interval = template?._config?.send_interval || 3;
      await delay(interval * 1000);
    }
  }

  // ------------------------------------------------
  // 浏览器自动化逻辑 (Puppeteer)
  // ------------------------------------------------
  async runBrowserBatch(sender: Electron.WebContents, tasks: any[], template: any) {
    try {
      const browser = await this.getBrowser();
      const pages = await browser.pages();
      const page = pages.length > 0 ? pages[0] : await browser.newPage();

      log.info('[Puppeteer] 正在打开 Gmail...');
      await page.goto('https://mail.google.com/mail/?authuser=0', { waitUntil: 'networkidle2', timeout: 60000 });

      const currentUrl = page.url();
      if (!currentUrl.includes('mail.google.com')) {
        log.warn('[Puppeteer] 最终网址未停在 Gmail 页面:', currentUrl);
        dialog.showMessageBox({
          type: 'warning',
          title: '需要登录',
          message: '检测到您的 Gmail 账户尚未登录，请先在浏览器任务中完成登录。',
          detail: '建议操作：点击“邮件配置”->"账户登录"后，并在弹出的浏览器中手动完成登录，然后重新启动发送任务。',
          buttons: ['确定']
        });
        tasks.filter(t => t.status === 'pending' || t.status === 'processing').forEach(t => {
          sender.send('batch-update', { id: t.id, status: 'failed', error: '请先登录账户' });
        });
        return;
      }

      for (const task of tasks) {
        if (task.status === 'success') continue;
        sender.send('batch-update', { id: task.id, status: 'processing' });
        try {
          log.info(`[Puppeteer] 正在为 ${task.receiver} 撰写邮件...`);
          const composeSelector = '.T-I-KE';
          const toSelector = 'input[aria-label="To recipients"]';
          
          let opened = false;
          let retries = 0;
          const maxRetries = 3;
          while (!opened && retries < maxRetries) {
            try {
              await page.waitForSelector(composeSelector, { visible: true, timeout: 10000 });
              await page.click(composeSelector);
              // 增加等待时间至 20 秒，并进行重试
              await page.waitForSelector(toSelector, { visible: true, timeout: 20000 });
              opened = true;
            } catch (e: any) {
              retries++;
              log.warn(`[Puppeteer] 等待收件人输入框失败 (第 ${retries} 次重试): ${e.message}`);
              if (retries >= maxRetries) throw e;
              await delay(2000); // 重试前稍作停顿
            }
          }

          const replaceVars = (str: string) => {
            if (!str) return '';
            return str.replace(/\{\{username\}\}/g, task.username || '');
          };

          const subject = replaceVars(template.subject);
          const body = replaceVars(template.content);

          await page.type(toSelector, task.receiver);
          await page.keyboard.press('Enter');
          const subjectSelector = 'input[name="subjectbox"]';
          await page.type(subjectSelector, subject);
          const bodySelector = 'div[aria-label="Message Body"]';
          await page.focus(bodySelector);
          await page.keyboard.type(body);
          const sendSelector = 'div[aria-label^="Send"]';
          // 使用 evaluate 强制点击，避免被遮挡导致 not clickable 错误
          await page.evaluate((selector) => {
            const btns = Array.from(document.querySelectorAll(selector)) as HTMLElement[];
            const visibleBtns = btns.filter(b => b.offsetWidth > 0 && b.offsetHeight > 0);
            if (visibleBtns.length > 0) {
              visibleBtns[visibleBtns.length - 1].click();
            }
          }, sendSelector);
          await delay(1500); // 稍微等待，检查是否出现警告
          
          try {
            // 检查 Send 按钮是否依然可见，通常如果出现警告(如外部联系人)，Send 按钮会依然保留
            const isSendBtnVisible = await page.evaluate((selector) => {
              const btns = Array.from(document.querySelectorAll(selector)) as HTMLElement[];
              const visibleBtns = btns.filter(b => b.offsetWidth > 0 && b.offsetHeight > 0);
              return visibleBtns.length > 0;
            }, sendSelector);

            if (isSendBtnVisible) {
              log.warn(`[Browser] 发现 Send 按钮仍然可见，可能遇到了发信警告，尝试再次点击 Send 按钮`);
              await page.evaluate((selector) => {
                const btns = Array.from(document.querySelectorAll(selector)) as HTMLElement[];
                const visibleBtns = btns.filter(b => b.offsetWidth > 0 && b.offsetHeight > 0);
                if (visibleBtns.length > 0) {
                  visibleBtns[visibleBtns.length - 1].click();
                }
              }, sendSelector);
              await delay(2000); // 再次延时等待发送完成
            } else {
              // 按钮消失，说明大概率成功，补足剩下的等待时间
              await delay(1500);
            }
          } catch (e: any) {
            log.warn(`[Browser] 检查警告状态时出错 (可忽略): ${e.message}`);
            await delay(1500);
          }

          sender.send('batch-update', { id: task.id, status: 'success' });
          log.info(`✅ [Browser] ${task.receiver} 发送成功`);
        } catch (err: any) {
          log.error(`❌ [Browser] ${task.receiver} 发送失败:`, err.message);
          sender.send('batch-update', { id: task.id, status: 'failed', error: err.message });
        }
        const interval = template?._config?.send_interval || 3;
        await delay(interval * 1000);
      }
      log.info('[Puppeteer] 所有浏览器发送任务处理完毕');

    } catch (error: any) {
      log.error('[Puppeteer) 错误:', error);
      tasks.filter(t => t.status === 'processing' || t.status === 'pending').forEach(t => {
        sender.send('batch-update', { id: t.id, status: 'failed', error: '浏览器错误: ' + error.message });
      });
    }
  }
}

export default homepageWindow;
