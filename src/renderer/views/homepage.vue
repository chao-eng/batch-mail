<template>
  <a-layout style="min-height: 100vh">
    <a-layout-sider v-model:collapsed="collapsed" collapsible theme="light" class="custom-sider">
      <div class="logo">
        <span v-if="!collapsed" class="logo-text">BatchMail</span>
        <span v-else class="logo-text">AM</span>
      </div>
      
      <a-menu v-model:selectedKeys="selectedKeys" theme="light" mode="inline">
        <a-menu-item key="config">
          <template #icon><SettingOutlined /></template>
          <span>邮件配置</span>
        </a-menu-item>
        
        <a-menu-item key="send">
          <template #icon><MailOutlined /></template>
          <span>发送邮件</span>
        </a-menu-item>

        <a-menu-item key="batch">
          <template #icon><FileExcelOutlined /></template>
          <span>批量发送</span>
        </a-menu-item>
      </a-menu>
    </a-layout-sider>

    <a-layout>
      <a-layout-header style="background: #fff; padding: 0 24px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 1px 4px rgba(0,21,41,0.08); z-index: 1;">
        <h2 style="margin: 0; font-size: 18px; font-weight: 600;">{{ pageTitle }}</h2>
      </a-layout-header>

      <a-layout-content style="margin: 24px 16px">
        <div style="padding: 24px; background: #fff; min-height: 360px; border-radius: 8px;">
          
          <div v-if="selectedKeys[0] === 'config'">
            <a-typography-title :level="4">SMTP 服务器配置</a-typography-title>
            <a-divider />
            <a-form 
              :model="formConfig" 
              :label-col="{ span: 4 }" 
              :wrapper-col="{ span: 14 }" 
              @finish="onSubmit"
            >
              <a-form-item label="服务器地址" name="smtp_server" :rules="[{ required: true, message: '请输入SMTP地址' }]">
                <a-input v-model:value="formConfig.smtp_server" placeholder="例如: smtp.gmail.com" />
              </a-form-item>
              
              <a-form-item label="端口" name="smtp_port" :rules="[{ required: true, message: '请输入端口' }]">
                <a-input-number v-model:value="formConfig.smtp_port" placeholder="465" style="width: 100%" />
              </a-form-item>
              
              <a-form-item label="邮箱账号" name="sender_email" :rules="[{ required: true, message: '请输入邮箱账号' }]">
                <a-input v-model:value="formConfig.sender_email" placeholder="your_email@gmail.com" />
              </a-form-item>
              
              <a-form-item label="授权码/密码" name="password" :rules="[{ required: true, message: '请输入授权码' }]">
                <a-input-password v-model:value="formConfig.password" placeholder="请输入应用专用密码" />
              </a-form-item>
              
              <a-form-item :wrapper-col="{ offset: 4, span: 14 }">
                <a-button type="primary" html-type="submit">保存配置</a-button>
                <a-button 
                  style="margin-left: 10px" 
                  @click="checkConfig" 
                  :loading="testingConfig"
                >
                  {{ testingConfig ? '测试中...' : '测试连接' }}
                </a-button>
              </a-form-item>
            </a-form>
          </div>

          <div v-else-if="selectedKeys[0] === 'send'">
             <a-typography-title :level="4">撰写新邮件</a-typography-title>
             <a-divider />
             <a-form layout="vertical" :model="mailForm">
               <a-form-item label="收件人">
                 <a-input v-model:value="mailForm.receiver" placeholder="请输入收件人邮箱，多个用逗号分隔" />
               </a-form-item>
               <a-form-item label="邮件主题">
                 <a-input v-model:value="mailForm.subject" placeholder="请输入邮件主题" />
               </a-form-item>
               <a-form-item label="邮件内容">
                 <a-textarea v-model:value="mailForm.content" placeholder="在此输入邮件正文..." :rows="8" />
               </a-form-item>
               <a-form-item>
                 <a-button type="primary" size="large" :loading="sending" @click="handleSendMail">
                    <template #icon><SendOutlined /></template>
                    {{ sending ? '发送中...' : '立即发送' }}
                 </a-button>
                 <a-button style="margin-left: 10px;" size="large" @click="clearMailForm">
                    重置
                 </a-button>
               </a-form-item>
             </a-form>
          </div>

          <div v-else-if="selectedKeys[0] === 'batch'">
            <div style="display: flex; justify-content: space-between; align-items: center;">
            <a-typography-title :level="4" style="margin: 0;"></a-typography-title>
            
            <a-space>
                <a-button 
                type="primary" 
                :disabled="tableData.length === 0 || isRunning"
                :loading="isRunning"
                @click="startBatch"
                >
                <template #icon><PlayCircleOutlined /></template>
                {{ isRunning ? '正在发送...' : '开始发送' }}
                </a-button>

                <a-button type="link" @click="downloadTemplate" >
                <template #icon><DownloadOutlined /></template>
                下载模板
                </a-button>
            </a-space>
            </div>
            <a-divider />

            <div v-if="tableData.length === 0" style="margin-bottom: 24px;">
            <a-upload-dragger
                name="file"
                :multiple="false"
                :before-upload="beforeUpload"
                accept=".xlsx,.xls"
                :showUploadList="false"
            >
                <p class="ant-upload-drag-icon"><InboxOutlined /></p>
                <p class="ant-upload-text">点击或拖拽 Excel 文件到这里</p>
                <p class="ant-upload-hint">支持 .xlsx / .xls，格式：收件人 | 主题 | 内容</p>
            </a-upload-dragger>
            </div>

            <a-table 
            v-else
            :columns="columns" 
            :data-source="tableData" 
            row-key="id"
            :pagination="{ pageSize: 10 }"
            size="middle"
            >
            <template #bodyCell="{ column, record }">
                <template v-if="column.key === 'status'">
                <a-tag v-if="record.status === 'pending'" color="default">待发送</a-tag>
                <a-tag v-else-if="record.status === 'processing'" color="blue">
                    <template #icon><SyncOutlined spin /></template> 发送中
                </a-tag>
                <a-tag v-else-if="record.status === 'success'" color="success">
                    <template #icon><CheckCircleOutlined /></template> 成功
                </a-tag>
                <a-tooltip v-else-if="record.status === 'failed'" :title="record.error">
                    <a-tag color="error" style="cursor: help;">
                    <template #icon><CloseCircleOutlined /></template> 失败
                    </a-tag>
                </a-tooltip>
                </template>
                <template v-if="column.key === 'error'">
                  <span v-if="record.status === 'failed'" style="color: #ff4d4f; font-size: 12px;">
                    {{ record.error }}
                  </span>
                </template>
            </template>
            </a-table>

            <div v-if="tableData.length > 0 && !isRunning" style="margin-top: 16px;">
                <a-button size="small" @click="clearTable">清空列表重新上传</a-button>
            </div>

        </div>

        </div>
      </a-layout-content>
      
      <a-layout-footer style="text-align: center">
        BatchMail Project ©2025 Created by bujic
        <a-divider type="vertical" />
        <GithubOutlined style="cursor: pointer; font-size: 16px" @click="openGithub" />
      </a-layout-footer>
    </a-layout>
  </a-layout>
</template>

<script setup lang="ts">
import { ref, computed, reactive, onMounted,onUnmounted } from 'vue';
import { message } from 'ant-design-vue';
import type { UploadProps } from 'ant-design-vue';
import { 
  SettingOutlined, 
  MailOutlined, 
  SendOutlined,
  FileExcelOutlined,
  InboxOutlined,
  DownloadOutlined,
  CloudUploadOutlined,
  GithubOutlined
} from '@ant-design/icons-vue';
import utils from "@utils/renderer";

const openGithub = () => {
  utils.openExternalLink("https://github.com/chao-eng/batch-mail");
};

// --- 全局状态 ---
const collapsed = ref<boolean>(false);
const selectedKeys = ref<string[]>(['config']);
const pageTitle = computed(() => {
  if (selectedKeys.value[0] === 'config') return '系统设置';
  if (selectedKeys.value[0] === 'send') return '邮件发送中心';
  if (selectedKeys.value[0] === 'batch') return '批量发送任务';
  return 'Dashboard';
});

// 获取 Electron API 的辅助函数
const getElectronApi = () => {
  // 增加简单的空值检查，防止在非 Electron 环境下报错
  const api = (window as any).homepageWindowAPI;
  if (!api) {
    console.warn("Electron API not found. Are you running in browser?");
    // 返回一个 mock 对象防止页面崩溃 (可选)
    return {
      setConfig: () => {},
      getConfig: async () => ({}),
      checkConfig: async () => ({ status: false, msg: 'API未连接' }),
      sendMail: async () => ({ status: false, msg: 'API未连接' })
    };
  }
  return api;
}

// --- 1. 配置文件相关逻辑 ---
const formConfig = reactive({
  smtp_server: "",
  smtp_port: "", // 绑定 input-number，这里如果是数字类型更好，但 String 兼容性更强
  sender_email: "",
  password: ""
});

// 新增：测试连接的 loading 状态
const testingConfig = ref(false);

// 初始化读取配置
onMounted(async () => {
  try {
    const config = await getElectronApi().getConfig("config");
    console.log("读取配置文件 config:", config);
    if (config) {
      formConfig.smtp_server = config.smtp_server || "";
      formConfig.smtp_port = config.smtp_port || "";
      formConfig.sender_email = config.sender_email || "";
      formConfig.password = config.password || "";
    }
  } catch (err) {
    console.error("读取配置失败", err);
  }
});

// 保存配置
const onSubmit = async (values: any) => {
  console.log("提交配置: ", values);
  try {
    // 确保把 reactive 对象转为普通对象再存，或者直接存 values
    await getElectronApi().setConfig("config", { ...formConfig });
    message.success("配置已保存");
  } catch (err) {
    console.error(err);
    message.error("保存失败");
  }
}

// 测试连接 (已修改：添加 Loading)
const checkConfig = async () => {
  // 简单校验
  if (!formConfig.smtp_server || !formConfig.password) {
    message.warning("请先填写完整的服务器信息");
    return;
  }

  testingConfig.value = true; // 1. 开启 loading
  
  try {
    const res = await getElectronApi().checkConfig({
        smtp_server: formConfig.smtp_server,
        smtp_port: formConfig.smtp_port,
        sender_email: formConfig.sender_email,
        password: formConfig.password
    });
    
    if (res.status) {
        message.success(res.msg);
    } else {
        message.error(res.msg);
    }
    console.log("Test Result:", res);
  } catch (err) {
    message.error("调用测试接口出错");
    console.error(err);
  } finally {
    testingConfig.value = false; // 2. 无论成功失败，关闭 loading
  }
}

// --- 2. 发送邮件相关逻辑 ---
const mailForm = reactive({
  receiver: "",
  subject: "",
  content: ""
});
const sending = ref(false);

const clearMailForm = () => {
  mailForm.receiver = "";
  mailForm.subject = "";
  mailForm.content = "";
};

const handleSendMail = async () => {
  if (!mailForm.receiver) return message.warning("请填写收件人邮箱");
  if (!mailForm.subject) return message.warning("请填写邮件主题");
  if (!formConfig.sender_email || !formConfig.password) {
    return message.error("请先在配置页完善配置");
  }

  sending.value = true;
  try {
    const payload = {
      smtp_server: formConfig.smtp_server,
      smtp_port: formConfig.smtp_port,
      sender_email: formConfig.sender_email,
      password: formConfig.password,
      to: mailForm.receiver,
      subject: mailForm.subject,
      text: mailForm.content,
    };

    const res = await getElectronApi().sendMail(payload);

    if (res.status) {
      message.success("邮件发送成功！");
    } else {
      message.error(`发送失败: ${res.msg}`);
    }
  } catch (error) {
    console.error(error);
    message.error("发送过程发生异常");
  } finally {
    sending.value = false;
  }
};

// --- 3. 批量发送相关逻辑 ---
// 定义接口
interface MailTask {
  id: string;
  receiver: string;
  subject: string;
  content: string;
  status: 'pending' | 'processing' | 'success' | 'failed';
  error?: string;
}

// 表格列定义
const columns = [
  { title: '收件人', dataIndex: 'receiver', key: 'receiver', width: 200 },
  { title: '主题', dataIndex: 'subject', key: 'subject' },
  { title: '状态', dataIndex: 'status', key: 'status', width: 120 },
  { title: '失败原因', dataIndex: 'error', key: 'error', ellipsis: true }
];

const tableData = ref<MailTask[]>([]);
const isRunning = ref(false);

// 1. 上传并解析 Excel
const beforeUpload = (file: File) => {
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const arrayBuffer = e.target?.result;
      // 调用解析接口
      const res = await getElectronApi().parseExcel({ fileData: arrayBuffer });
      
      if (res.status) {
        tableData.value = res.data; // 将解析后的数据赋值给表格
        message.success(res.msg);
      } else {
        message.error(res.msg);
      }
    } catch (err) {
      console.error(err);
      message.error('解析失败');
    }
  };
  reader.readAsArrayBuffer(file);
  return false; // 阻止默认上传
};

// 2. 开始批量发送
const startBatch = async () => {
  if (tableData.value.length === 0) return;
  
  // 过滤出未成功的任务
  const pendingTasks = tableData.value.filter(t => t.status !== 'success');
  if (pendingTasks.length === 0) {
    message.info('所有任务已完成，无需发送');
    return;
  }

  isRunning.value = true;
  try {
    // 这里把整个列表传给后台（或者只传 pending 的，看你后台逻辑）
    // 后台拿到列表后会立即返回 "启动成功"，不会等待所有发完
    const plainTasks = JSON.parse(JSON.stringify(tableData.value));
    const res = await getElectronApi().startBatchTasks(plainTasks);
    if (res.status) {
      message.success('任务已启动，请保持网络畅通');
    } else {
      message.error(res.msg);
      isRunning.value = false;
    }
  } catch (err) {
    console.error(err);
    isRunning.value = false;
  }
};

// 3. 监听后台发来的进度事件
onMounted(() => {
  const api = getElectronApi();
  if (api) {
    // 监听 batch-update 事件
    api.onBatchUpdate((event: any, updateData: any) => {
      // updateData 结构: { id: '...', status: '...', error: '...' }
      console.log('收到进度更新:', updateData);

      // 在表格数据中找到对应的那一行，更新它的状态
      const task = tableData.value.find(t => t.id === updateData.id);
      if (task) {
        task.status = updateData.status;
        if (updateData.error) {
          task.error = updateData.error;
        }
      }

      // 如果所有任务都不是 processing 或 pending，说明全部结束了
      // 这里的逻辑可以简单判断，也可以等后台发专门的 finish 事件
      const hasRunning = tableData.value.some(t => t.status === 'processing');
      // 注意：这只是个简单的本地判断，也可以不自动停止 loading，由用户手动或等全部完成后自动停止
      if (!hasRunning && updateData.status !== 'processing') {
         // 检查是否还有 pending 的，如果没有了，就停止 loading
         const hasPending = tableData.value.some(t => t.status === 'pending');
         if (!hasPending) {
             isRunning.value = false;
             message.success('批量发送队列执行完毕');
         }
      }
    });
  }
});

// 组件卸载时移除监听
onUnmounted(() => {
  const api = getElectronApi();
  if (api && api.removeBatchUpdateListener) {
    api.removeBatchUpdateListener();
  }
});

const clearTable = () => {
  tableData.value = [];
  isRunning.value = false;
};
// 下载模板
const downloadTemplate = async () => {
  try {
    const res = await getElectronApi().downloadTemplate();
    if (res.status) {
      message.success(res.msg);
    } else {
      if (res.msg !== '取消下载') {
        message.error(res.msg);
      }
    }
  } catch (err) {
    console.error(err);
    message.error("调用下载接口失败");
  }
};

</script>

<style scoped>
/* Logo 样式 */
.logo {
  height: 64px;
  display: flex;
  align-items: center;
  justify-content: center;
  background: rgba(0, 0, 0, 0.02);
  margin-bottom: 16px;
  overflow: hidden;
  transition: all 0.2s;
}

.logo-text {
  font-size: 18px;
  font-weight: bold;
  color: #1890ff;
  white-space: nowrap;
}

/* 侧边栏阴影 */
.custom-sider {
  box-shadow: 2px 0 8px 0 rgba(29, 35, 41, 0.05);
  z-index: 2;
}

.ant-layout-content {
  transition: all 0.2s;
}
</style>