<script lang="ts">
  import { onMount, tick } from 'svelte';
  import Header from './components/Header.svelte';
  import SettingsModal from './components/SettingsModal.svelte';
  import ScriptDrawer from './components/ScriptDrawer.svelte';
  import ExecutionCard from './components/ExecutionCard.svelte';
  import ChatInput from './components/ChatInput.svelte';
  import { bridge, type WorkbookInfo, type ScriptItem, type VbaExecutionData } from './services/bridge';
  import { callLlmStream, extractVbaCode, parseStreamOutput, hasValidVbaTask } from './services/llm';

  interface ChatMessage {
    id: string;
    role: 'user' | 'assistant';
    content: string; // 纯文本内容/简述
    prompt?: string;
    streamVbaCode?: string; // 流式生成的代码
    execution?: VbaExecutionData | null;
    isExecuting?: boolean;
  }

  let workbook: WorkbookInfo | null = null;
  let messages: ChatMessage[] = [];
  let isProcessing = false;
  let messagesContainer: HTMLElement;

  let showSettings = false;
  let showScripts = false;
  let scriptDrawerRef: any;

  onMount(async () => {
    // 监听 C# 宿主推送的工作簿激活变更
    bridge.onWorkbookChange((info) => {
      workbook = info;
    });

    // 初次获取工作簿信息
    await refreshWorkbookInfo();

    // 定时轮询，自动感知工作簿新建/打开/切换
    const timer = setInterval(() => {
      refreshWorkbookInfo();
    }, 2500);

    // 默认添加欢迎引导语
    messages = [
      {
        id: 'msg_welcome',
        role: 'assistant',
        content: '你好！我是你的 Excel AI 助手。已自动连接当前活动工作簿。\n你可以用自然语言向我下达任何数据分析、表格整理或自动化操作指令。每次运行前会自动保存整本物理快照，随时支持一键恢复。',
      },
    ];

    return () => clearInterval(timer);
  });

  async function refreshWorkbookInfo() {
    const res = await bridge.send<WorkbookInfo>('get_workbook_info');
    if (res.ok && res.data) {
      workbook = res.data;
    }
  }

  async function scrollToBottom() {
    await tick();
    if (messagesContainer) {
      messagesContainer.scrollTop = messagesContainer.scrollHeight;
    }
  }

  // 核心智能交互与全自动执行闭环
  async function handleSend(text: string) {
    if (!text.trim() || isProcessing) return;

    isProcessing = true;

    // 1. 添加用户消息
    const userMsgId = 'user_' + Date.now();
    messages = [
      ...messages,
      {
        id: userMsgId,
        role: 'user',
        content: text,
      },
    ];
    await scrollToBottom();

    // 2. 占位助理消息
    const assistantMsgId = 'ai_' + Date.now();
    const assistantMsg: ChatMessage = {
      id: assistantMsgId,
      role: 'assistant',
      content: '正在思考...',
      prompt: text,
      streamVbaCode: '',
      isExecuting: true,
      execution: null,
    };
    messages = [...messages, assistantMsg];
    await scrollToBottom();

    try {
      // 3. 构建历史对话上下文 (排除当前消息)
      const chatHistory = messages
        .filter((m) => m.id !== assistantMsgId && m.id !== userMsgId && m.id !== 'msg_welcome')
        .map((m) => ({ role: m.role, content: m.content }));

      const wbName = workbook?.name || '';
      const sheets = workbook?.sheets || [];

      // 4. 调用大模型流式生成
      const rawResponse = await callLlmStream(
        text,
        wbName,
        sheets,
        (partialText) => {
          const parsed = parseStreamOutput(partialText);
          if (parsed.hasCode) {
            // 操作指令：气泡只展示一两句简报，代码流入折叠卡片
            assistantMsg.content = parsed.explanation || '正在执行自动化指令...';
            assistantMsg.streamVbaCode = parsed.vbaCode;
          } else {
            // 普通对话/闲聊/问答：气泡直接流式展示自然语言文本
            assistantMsg.content = partialText;
            assistantMsg.streamVbaCode = '';
          }
          messages = [...messages];
        },
        chatHistory
      );

      // 5. 严格验证是否包含真实合法的 VBA 任务 (必须同时包含 Sub 和 End Sub)
      const vbaCode = extractVbaCode(rawResponse);
      const isTask = hasValidVbaTask(vbaCode);

      if (isTask) {
        const finalParsed = parseStreamOutput(rawResponse);
        assistantMsg.content = finalParsed.explanation || '已生成操作方案并执行';
        assistantMsg.streamVbaCode = vbaCode;

        // 6. 调用 C# 原生宿主执行 (自动快照 + 注入运行 + 瞬时清理)
        const execRes = await bridge.send<VbaExecutionData>('execute_vba', {
          code: vbaCode,
          prompt: text,
        });

        if (execRes.ok && execRes.data) {
          assistantMsg.execution = execRes.data;
        } else {
          assistantMsg.execution = {
            summary: execRes.error || '执行遇到阻断',
            error: execRes.error || 'COM 调用失败',
            elapsedMs: 0,
            vbaCode,
          };
        }

        // 刷新工作簿与快照时间轴
        await refreshWorkbookInfo();
      } else {
        // 纯文本回复 (问答/咨询/闲聊/功能解释)
        // 绝对不调用 execute_vba，绝对不挂起任何报错卡片
        assistantMsg.content = rawResponse.replace(/```(?:vba|vb)?[\s\S]*?```/gi, '').trim();
        assistantMsg.streamVbaCode = '';
        assistantMsg.execution = null;
      }
    } catch (err: any) {
      // 网络或接口本身异常
      assistantMsg.content = assistantMsg.content || '处理遇到异常';
      assistantMsg.execution = {
        summary: '操作中断：' + (err.message || '未知异常'),
        error: err.stack || err.message,
        elapsedMs: 0,
        vbaCode: assistantMsg.streamVbaCode || '',
      };
    } finally {
      assistantMsg.isExecuting = false;
      messages = [...messages];
      isProcessing = false;
      await scrollToBottom();
    }
  }

  // 从“我的脚本”直接运行
  async function handleRunScript(script: ScriptItem) {
    if (isProcessing) return;
    isProcessing = true;

    const userMsgId = 'user_script_' + Date.now();
    messages = [
      ...messages,
      {
        id: userMsgId,
        role: 'user',
        content: `运行脚本: 【${script.name}】`,
      },
    ];

    const assistantMsgId = 'ai_script_' + Date.now();
    const assistantMsg: ChatMessage = {
      id: assistantMsgId,
      role: 'assistant',
      content: `准备运行本地脚本: ${script.name}`,
      prompt: `运行脚本: ${script.name}`,
      isExecuting: true,
      execution: null,
    };
    messages = [...messages, assistantMsg];
    await scrollToBottom();

    try {
      const execRes = await bridge.send<VbaExecutionData>('execute_vba', {
        code: script.code,
        prompt: `运行脚本: ${script.name}`,
      });

      if (execRes.ok && execRes.data) {
        assistantMsg.execution = execRes.data;
      } else {
        assistantMsg.execution = {
          summary: execRes.error || '脚本执行失败',
          error: execRes.error,
          elapsedMs: 0,
          vbaCode: script.code,
        };
      }
      await refreshWorkbookInfo();
    } catch (e: any) {
      assistantMsg.execution = {
        summary: '运行失败: ' + e.message,
        error: e.message,
        elapsedMs: 0,
        vbaCode: script.code,
      };
    } finally {
      assistantMsg.isExecuting = false;
      messages = [...messages];
      isProcessing = false;
      await scrollToBottom();
    }
  }
</script>

<div class="app-layout">
  <!-- 顶部状态栏 -->
  <Header
    {workbook}
    onOpenSettings={() => (showSettings = true)}
    onOpenScripts={() => (showScripts = true)}
    onRefresh={refreshWorkbookInfo}
  />

  <!-- 消息流区域 -->
  <main class="chat-area" bind:this={messagesContainer}>
    {#each messages as msg (msg.id)}
      {#if msg.role === 'user'}
        <div class="message-row user-row">
          <div class="user-bubble">{msg.content}</div>
        </div>
      {:else}
        <div class="message-row ai-row">
          {#if msg.content}
            <div class="ai-text">{msg.content}</div>
          {/if}

          <!-- 全自动执行结果卡片 (仅在确有 VBA 自动化指令时展示，纯自然语言对话时不显示) -->
          {#if (msg.isExecuting && msg.streamVbaCode) || msg.execution}
            <ExecutionCard
              prompt={msg.prompt || ''}
              execution={msg.execution}
              streamCode={msg.streamVbaCode || ''}
              isExecuting={msg.isExecuting || false}
              allSnapshots={workbook?.snapshots || []}
              onSaveScriptSuccess={() => scriptDrawerRef?.refreshScripts()}
            />
          {/if}
        </div>
      {/if}
    {/each}
  </main>

  <!-- 底部紧凑输入栏 -->
  <ChatInput disabled={isProcessing} onSend={handleSend} />

  <!-- 设置弹窗 -->
  {#if showSettings}
    <SettingsModal onClose={() => (showSettings = false)} />
  {/if}

  <!-- “我的脚本”抽屉 -->
  <ScriptDrawer
    bind:this={scriptDrawerRef}
    isOpen={showScripts}
    onClose={() => (showScripts = false)}
    onRunScript={handleRunScript}
  />
</div>

<style>
  .app-layout {
    display: flex;
    flex-direction: column;
    height: 100vh;
    width: 100vw;
    background: var(--office-bg);
    overflow: hidden;
  }

  .chat-area {
    flex: 1;
    overflow-y: auto;
    padding: 12px;
    display: flex;
    flex-direction: column;
    gap: 12px;
  }

  .message-row {
    display: flex;
    flex-direction: column;
    max-width: 100%;
  }

  .user-row {
    align-items: flex-end;
  }

  .user-bubble {
    background: var(--excel-green);
    color: white;
    padding: 7px 12px;
    border-radius: 6px 6px 1px 6px;
    font-size: 12px;
    line-height: 1.4;
    max-width: 85%;
    word-break: break-all;
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.08);
  }

  .ai-row {
    align-items: flex-start;
  }

  .ai-text {
    background: white;
    padding: 8px 12px;
    border-radius: 6px 6px 6px 1px;
    font-size: 12px;
    line-height: 1.5;
    border: 1px solid var(--office-border);
    max-width: 95%;
    white-space: pre-wrap;
    word-break: break-all;
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04);
  }
</style>
