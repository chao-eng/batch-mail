import path from "path";
import WindowBase from "../window-base";
import configManager from "./config-manager";
import * as XLSX from 'xlsx';
import nodemailer, { Transporter, SendMailOptions, SentMessageInfo } from 'nodemailer';
import * as fs from 'fs'; // 记得引入 fs
import { dialog } from 'electron';

// 定义配置接口 (方便你以后在业务代码中复用)
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



import appState from "../../app-state";

class homepageWindow extends WindowBase {
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
    //获取配置文件
    this.registerIpcHandleHandler('getConfig', (event, key) => configManager.get(key))
    //保存配置文件
    this.registerIpcHandleHandler('setConfig', async (event, key, value) => {
      await configManager.set(key, value);
      console.log("setConfig", key, value);
      console.log("保存目录: ", await configManager.getPath());
    });
    //测试连接
    this.registerIpcHandleHandler('checkConfig', async (event, rawConfig) => {
      console.log("checkConfig", rawConfig);
      try {
        const config: SmtpConfig = {
          host: rawConfig.smtp_server,
          port: rawConfig.smtp_port,            // 465 (SSL) 或 587 (TLS)
          secure: true,         // 端口465设为true，端口587设为false
          auth: {
            user: rawConfig.sender_email,  // 替换
            pass: rawConfig.password, // 替换为16位应用专用密码
          },
        };
        console.log('🚀 开始测试 SMTP 连接 (TypeScript Mode)...');
        // ---  创建 Transporter 对象 ---
        const transporter: Transporter = nodemailer.createTransport(config);

        //  3. 验证连接 (Verify) ---
        // 这一步专门用于检查密码和服务器配置是否正确
        const verifySuccess: boolean = await transporter.verify();

        if (verifySuccess) {
          console.log('✅ [连接成功] 服务器配置正确，身份验证通过！');
          return { status: true, msg: "服务器配置正确，身份验证通过" }
        }
      } catch (error: any) {
        console.error('❌ [发生错误]');

        // 简单的类型守卫或直接读取 message
        if (error instanceof Error) {
          console.error(`   错误信息: ${error.message}`);

          if ((error as any).code === 'EAUTH') {
            console.warn('   👉 提示: 认证失败。请检查邮箱账号或应用专用密码。');
            return { status: false, msg: "👉 提示: 认证失败。请检查邮箱账号或应用专用密码。" }
          }
        } else {
          console.error('   未知错误:', error);
        }
        return { status: false, msg: error.message }
      }

    })
    //发送邮件
    this.registerIpcHandleHandler('sendMail', async (event, params) => {
      console.log('发送邮件:', params)
      const { smtp_server, smtp_port, sender_email, password, to, cc, subject, text, html, attachments } = params;

      // 1. 创建 Transporter
      const transporter = nodemailer.createTransport({
        host: smtp_server,
        port: parseInt(smtp_port),
        secure: parseInt(smtp_port) === 465, // 简单判断 SSL
        auth: {
          user: sender_email,
          pass: password,
        },
      });

      try {
        // 2. 发送邮件
        const info = await transporter.sendMail({
          from: sender_email,
          to: to,
          cc: cc, // 抄送
          subject: subject,
          text: text, // 纯文本
          attachments: attachments, // 附件数组
          html: html || text // 优先使用 html，兼容旧代码
        });

        return { status: true, msg: '发送成功', id: info.messageId };
      } catch (error: any) {
        return { status: false, msg: error.message };
      }
    })
    // ------------------------------------------------
    //  解析 Excel 并返回列表
    // ------------------------------------------------
    this.registerIpcHandleHandler('parseExcel', async (event, payload) => {
      const { fileData } = payload;
      console.log('📄 正在解析 Excel...');

      try {
        const workbook = XLSX.read(fileData, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const rows: any[][] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

        const tasks: any[] = [];

        // 从第 1 行开始遍历（跳过表头）
        rows.forEach((row, index) => {
          if (index === 0) return;

          const receiver = row[0];
          const username = row[1] || ''; // 第二列为用户名

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
    // ------------------------------------------------
    //  执行批量任务 (流式反馈)
    // ------------------------------------------------
    this.registerIpcHandleHandler('startBatchTasks', async (event, tasks: any[]) => {
      console.log(`🚀 开始执行 ${tasks.length} 个任务`);

      // 1. 获取 SMTP 配置
      const config = await configManager.get('config');
      if (!config?.smtp_server) {
        return { status: false, msg: 'SMTP 配置未找到' };
      }

      // 2. 获取邮件模板配置
      const template = await configManager.get('email_template');
      if (!template || !template.subject) {
        return { status: false, msg: '邮件模板未配置（请先设置主题和内容）' };
      }

      // 3. 初始化 Transporter
      const transporter = nodemailer.createTransport({
        host: config.smtp_server,
        port: parseInt(config.smtp_port),
        secure: parseInt(config.smtp_port) === 465,
        auth: { user: config.sender_email, pass: config.password },
      });

      // 4. 异步开始循环
      // 将配置合并到 template 中传递给 processQueue，或者直接传 config
      const templateWithConfig = { ...template, _config: config };
      this.processQueue(event.sender, transporter, tasks, config.sender_email, templateWithConfig, config.cc_emails);

      return { status: true, msg: '任务队列已启动' };
    });
    //下载模板
    // --- 新增：下载模板 ---
    this.registerIpcHandleHandler('downloadTemplate', async () => {
      // 1. 定义表头数据
      const headers = [['收件人', '用户名']];

      // 2. 创建 Workbook 和 Sheet
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet(headers);

      // (可选) 设置列宽
      worksheet['!cols'] = [{ wch: 30 }, { wch: 20 }];

      XLSX.utils.book_append_sheet(workbook, worksheet, '批量发送模板');

      // 3. 生成 Buffer
      const buffer = XLSX.write(workbook, { type: 'buffer' });

      // 4. 弹出保存对话框
      const { canceled, filePath } = await dialog.showSaveDialog({
        title: '保存 Excel 模板',
        defaultPath: '邮件批量发送模板.xlsx',
        filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
      });

      if (canceled || !filePath) {
        return { status: false, msg: '取消下载' };
      }

      try {
        // 5. 写入文件
        fs.writeFileSync(filePath, buffer);
        return { status: true, msg: '模板已保存' };
      } catch (error: any) {
        return { status: false, msg: '保存失败: ' + error.message };
      }
    });

    // --- 新增：导出发送结果 ---
    this.registerIpcHandleHandler('exportResults', async (event, data: any[]) => {
      // 1. 准备数据
      const exportData = data.map(item => ({
        '收件人': item.receiver,
        '用户名': item.username,
        '状态': item.status === 'success' ? '成功' : (item.status === 'failed' ? '失败' : '待发送'),
        '失败原因': item.error || ''
      }));

      // 2. 创建 Workbook 和 Sheet
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.json_to_sheet(exportData);

      // 设置列宽
      worksheet['!cols'] = [{ wch: 30 }, { wch: 20 }, { wch: 15 }, { wch: 40 }];

      XLSX.utils.book_append_sheet(workbook, worksheet, '发送结果');

      // 3. 生成 Buffer
      const buffer = XLSX.write(workbook, { type: 'buffer' });

      // 4. 弹出保存对话框
      const { canceled, filePath } = await dialog.showSaveDialog({
        title: '导出发送结果',
        defaultPath: `发送结果_${new Date().toISOString().split('T')[0]}.xlsx`,
        filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
      });

      if (canceled || !filePath) {
        return { status: false, msg: '取消导出' };
      }

      try {
        // 5. 写入文件
        fs.writeFileSync(filePath, buffer);
        return { status: true, msg: '结果已导出' };
      } catch (error: any) {
        return { status: false, msg: '导出保存失败: ' + error.message };
      }
    });

  }

  // ------------------------------------------------
  // 辅助方法：处理队列
  // ------------------------------------------------
  async processQueue(sender: Electron.WebContents, transporter: any, tasks: any[], fromEmail: string, template: any, ccEmails?: string) {
    for (const task of tasks) {
      if (task.status === 'success') continue;

      sender.send('batch-update', { id: task.id, status: 'processing' });

      try {
        // 自动替换变量 {{username}}
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
        console.log(`✅ [${task.receiver}] 发送成功`);

      } catch (err: any) {
        sender.send('batch-update', { id: task.id, status: 'failed', error: err.message });
        console.error(`❌ [${task.receiver}] 发送失败: ${err.message}`);
      }

      // 使用配置的发送间隔，默认 3s
      const interval = template?._config?.send_interval || 3;
      await delay(interval * 1000);
    }
  }


}



export default homepageWindow;
