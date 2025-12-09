# BatchMail - 批量邮件发送工具

BatchMail 是一款基于 **Electron** + **Vue 3** + **TypeScript** 开发的现代化桌面端批量邮件发送应用。它旨在帮助用户通过简单的配置和 Excel 导入，轻松实现邮件的批量发送与管理。

![Main UI](./screenshot/main.png)

## ✨ 主要特性

- **SMTP 配置管理**：支持自定义 SMTP 服务器（如 Gmail, QQ 邮箱, 网易邮箱等），并提供连接测试功能。
- **单条发送**：提供便捷的单条邮件发送界面，支持多收件人。
- **批量发送**：
  - 支持导入 Excel (`.xlsx`, `.xls`) 文件批量创建发送任务。
  - 提供标准 Excel 模板下载。
  - 实时显示发送进度、成功/失败状态。
  - 详细的失败原因展示。
- **现代化 UI**：基于 Ant Design Vue 开发，界面简洁美观，操作直观。
- **跨平台**：支持 Windows, macOS 和 Linux。

## 🛠️ 技术栈

- **Core**: [Electron](https://www.electronjs.org/)
- **Frontend**: [Vue 3](https://vuejs.org/) + [TypeScript](https://www.typescriptlang.org/) + [Vite](https://vitejs.dev/)
- **UI Framework**: [Ant Design Vue](https://www.antdv.com/)
- **Mail Engine**: [Nodemailer](https://nodemailer.com/)
- **Excel Processing**: [SheetJS (xlsx)](https://sheetjs.com/)

## 🚀 快速开始

### 1. 环境准备

确保您的电脑上已安装 [Node.js](https://nodejs.org/) (推荐 v16+) 和 [Yarn](https://yarnpkg.com/)。

### 2. 克隆项目

```bash
git clone https://github.com/your-repo/batch-mail.git
cd batch-mail
```

### 3. 安装依赖

```bash
yarn install
```

> 国内用户建议配置 Yarn 淘宝源以加速下载。

### 4. 启动开发环境

```bash
yarn run dev
```

启动后，应用将以开发模式运行，支持热重载。

## 📦 打包构建

本项目使用 Electron Forge 进行打包。

```bash
# 构建当前平台的应用
yarn run build

# 构建 Windows 平台应用
yarn run build:win32  # 32位
yarn run build:win64  # 64位

# 构建 macOS 应用
yarn run build:mac

# 构建 Linux 应用
yarn run build:linux

# 生成 Windows 安装包 (NSIS)
# 需确保 setup/NSIS 环境已配置
yarn run build:nsis-win32
yarn run build:nsis-win64
```

构建产物将位于 `out` 目录下。

## 📖 使用指南

1.  **配置邮箱**：
    *   进入“邮件配置”页面。
    *   填写 SMTP 服务器地址（如 `smtp.qq.com`）、端口（如 `465`）。
    *   填写发件人邮箱和授权码（注意：通常不是登录密码，而是邮箱服务商提供的 SMTP 授权码）。
    *   点击“测试连接”确保配置正确，然后点击“保存配置”。

2.  **发送单条邮件**：
    *   进入“发送邮件”页面，填写收件人、主题和内容即可发送。

3.  **批量发送**：
    *   进入“批量发送”页面。
    *   点击“下载模板”获取 Excel 模板文件。
    *   在 Excel 中按格式填入收件人、主题和内容。
    *   拖拽或点击上传 Excel 文件。
    *   点击“开始发送”，系统将自动处理队列。

## 📄 许可证

[MIT License](LICENSE)