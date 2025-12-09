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
      const { smtp_server, smtp_port, sender_email, password, to, subject, text, attachments } = params;

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
          subject: subject,
          // text: text, // 纯文本
          attachments: attachments, // 附件数组
          html: text // 如果你想支持 html，可以用这个字段
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

        // 从第 0 行开始遍历（假设没有表头，或者你在前端做校验）
        rows.forEach((row, index) => {
          // --- 修改点 Start: 跳过表头 ---
          if (index === 0) {
            return; // 这里的 return 相当于 for循环里的 continue，跳过本次回调
          }
          const receiver = row[0];
          // 简单验证有效性
          if (receiver && typeof receiver === 'string' && receiver.includes('@')) {
            // 解析附件列 (第四列, index 3)
            let attachments: string[] = [];
            if (row[3]) {
              const raw = String(row[3]);
              attachments = raw.split(';').map(p => p.trim()).filter(p => p.length > 0);
            }

            tasks.push({
              id: `task-${Date.now()}-${index}`, // 生成唯一ID
              receiver: receiver,
              subject: row[1] || '无主题',
              content: row[2] || '',
              attachments: attachments,
              status: 'pending', // 初始状态：待发送
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

      // 1. 获取配置
      const config = await configManager.get('config');
      if (!config?.smtp_server) {
        return { status: false, msg: 'SMTP 配置未找到' };
      }

      // 2. 初始化 Transporter
      const transporter = nodemailer.createTransport({
        host: config.smtp_server,
        port: parseInt(config.smtp_port),
        secure: parseInt(config.smtp_port) === 465,
        auth: { user: config.sender_email, pass: config.password },
      });

      // 3. 异步开始循环 (不 await 整个循环，直接让 handle 返回，告诉前端“任务已启动”)
      // 注意：这里我们不阻塞主线程返回，而是开启一个异步过程
      this.processQueue(event.sender, transporter, tasks, config.sender_email);

      return { status: true, msg: '任务队列已启动' };
    });
    //下载模板
    // --- 新增：下载模板 ---
    this.registerIpcHandleHandler('downloadTemplate', async () => {
      // 1. 定义表头数据
      const headers = [['收件人', '主题', '邮件内容', '附件地址(多个用;分隔)']];

      // 2. 创建 Workbook 和 Sheet
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet(headers);

      // (可选) 设置列宽，让模板好看点
      worksheet['!cols'] = [{ wch: 25 }, { wch: 20 }, { wch: 40 }, { wch: 40 }];

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

  }

  // ------------------------------------------------
  // 辅助方法：处理队列
  // ------------------------------------------------
  async processQueue(sender: Electron.WebContents, transporter: any, tasks: any[], fromEmail: string) {
    for (const task of tasks) {
      // 如果前端传过来的列表里包含非 pending 的（比如之前成功的），跳过
      if (task.status === 'success') continue;

      // 1. 通知前端：正在处理
      sender.send('batch-update', { id: task.id, status: 'processing' });

      try {
        // 构建附件数组
        const attachments = task.attachments ? task.attachments.map((p: string) => ({ path: p })) : [];

        await transporter.sendMail({
          from: fromEmail,
          to: task.receiver,
          subject: task.subject,
          //text: task.content,
          html: task.content,
          attachments: attachments
        });

        // 2. 通知前端：成功
        sender.send('batch-update', { id: task.id, status: 'success' });
        console.log(`✅ [${task.receiver}] 发送成功`);

      } catch (err: any) {
        // 3. 通知前端：失败
        sender.send('batch-update', { id: task.id, status: 'failed', error: err.message });
        console.error(`❌ [${task.receiver}] 发送失败: ${err.message}`);
      }

      // 4. 延时 (防封号)
      await delay(1500);
    }

    // 全部结束可以发个完成事件 (可选)
    // sender.send('batch-finished');
  }


}



export default homepageWindow;
