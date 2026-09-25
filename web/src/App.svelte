<script lang="ts">
  import { onMount, tick } from 'svelte';
  import Header from './components/Header.svelte';
  import SettingsModal from './components/SettingsModal.svelte';
  import ScriptDrawer from './components/ScriptDrawer.svelte';
  import ExecutionCard from './components/ExecutionCard.svelte';
  import ChatInput from './components/ChatInput.svelte';
  import { bridge, type WorkbookInfo, type ScriptItem, type VbaExecutionData } from './services/bridge';
  import {
    callLlmStream,
    extractVbaCode,
    parseStreamOutput,
    detectIntent,
    verifyExecutionResult,
  } from './services/llm';

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
        content:
          '你好！我是你的 Excel AI 助手。已自动绑定当前目标工作簿。\n你可以用自然语言下达表格操作、数据汇总或进行 Excel 技巧咨询。普通问答直接解答；操作指令在每次运行前自动保存整本物理副本，支持一键回滚。',
      },
    ];

    return () => clearInterval(timer);
  });

  async function refreshWorkbookInfo() {
    const res = await bridge.send<WorkbookInfo>('get_workbook_info', {
      targetWorkbookName: workbook?.name || '',
      targetWorkbookFullName: workbook?.fullName || '',
    });
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

  // 核心交互：通道分离、严格代码解析、目标锁定与写后核验
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

    // 2. 意图通道路由识别 (CHAT | AUTOMATION | AMBIGUOUS)
    const intent = detectIntent(text);

    // 场景 A: 意图模糊不清时，只问简短澄清问题，绝不执行、绝不注入、绝不创建快照
    if (intent === 'AMBIGUOUS') {
      const assistantMsgId = 'ai_' + Date.now();
      messages = [
        ...messages,
        {
          id: assistantMsgId,
          role: 'assistant',
          content: '请问您是希望了解该内容，还是需要在当前工作簿中执行具体操作？如果是操作表格，请简要说明您的具体需求。',
          streamVbaCode: '',
          execution: null,
          isExecuting: false,
        },
      ];
      isProcessing = false;
      await scrollToBottom();
      return;
    }

    // 场景 B & C: 明确的 CHAT 或 AUTOMATION
    const assistantMsgId = 'ai_' + Date.now();
    const assistantMsg: ChatMessage = {
      id: assistantMsgId,
      role: 'assistant',
      content: intent === 'CHAT' ? '正在思考解答...' : '正在准备自动化方案...',
      prompt: text,
      streamVbaCode: '',
      isExecuting: true,
      execution: null,
    };
    messages = [...messages, assistantMsg];
    await scrollToBottom();

    // 关键：在任务发起时刻，明确锁定目标工作簿身份凭据 (由宿主保持并校验)
    const targetWbName = workbook?.name || '';
    const targetWbFullName = workbook?.fullName || '';
    const sheets = workbook?.sheets || [];
    const activeSheet = workbook?.activeSheetName || '';
    const usedRange = workbook?.usedRangeAddress || '';

    try {
      const chatHistory = messages
        .filter((m) => m.id !== assistantMsgId && m.id !== userMsgId && m.id !== 'msg_welcome')
        .map((m) => ({ role: m.role, content: m.content }));

      // 调用大模型流式生成 (根据 intent 选用对应的系统提示词)
      const rawResponse = await callLlmStream(
        text,
        targetWbName,
        sheets,
        activeSheet,
        usedRange,
        intent,
        (partialText) => {
          if (intent === 'AUTOMATION') {
            const parsed = parseStreamOutput(partialText);
            if (parsed.hasCode) {
              assistantMsg.content = parsed.explanation || '正在执行自动化指令...';
              assistantMsg.streamVbaCode = parsed.vbaCode;
            } else {
              assistantMsg.content = partialText;
              assistantMsg.streamVbaCode = '';
            }
          } else {
            // CHAT 通道：流式显示纯文本
            assistantMsg.content = partialText;
            assistantMsg.streamVbaCode = '';
          }
          messages = [...messages];
        },
        chatHistory
      );

      // 处理 CHAT 通道：确保绝对零 VBA 调用、零快照
      if (intent === 'CHAT') {
        assistantMsg.content = rawResponse.replace(/```(?:vba|vb)?[\s\S]*?```/gi, '').trim() || rawResponse;
        assistantMsg.streamVbaCode = '';
        assistantMsg.execution = null;
        assistantMsg.isExecuting = false;
        messages = [...messages];
        return;
      }

      // 处理 AUTOMATION 通道：严格提取与校验 VBA 代码
      const extracted = extractVbaCode(rawResponse);

      if (extracted.error || !extracted.code) {
        // 响应被截断、格式不合法或缺少代码时，停止执行并显示真实原因
        const errorReason = extracted.error || '模型未返回有效的 VBA 过程代码';
        const parsed = parseStreamOutput(rawResponse);
        assistantMsg.content = parsed.explanation ? `${parsed.explanation}\n（${errorReason}）` : errorReason;
        assistantMsg.streamVbaCode = extracted.code || '';
        assistantMsg.execution = {
          summary: errorReason,
          error: errorReason,
          elapsedMs: 0,
          vbaCode: extracted.code || '',
        };
        assistantMsg.isExecuting = false;
        messages = [...messages];
        return;
      }

      // 提取成功且结构完整：调用宿主执行 (显式绑定目标工作簿)
      const parsed = parseStreamOutput(rawResponse);
      assistantMsg.streamVbaCode = extracted.code;

      const execRes = await bridge.send<VbaExecutionData>('execute_vba', {
        code: extracted.code,
        prompt: text,
        targetWorkbookName: targetWbName,
        targetWorkbookFullName: targetWbFullName,
      });

      if (execRes.ok && execRes.data) {
        // 执行后核验：比对写后读回与用户实际诉求
        const verification = verifyExecutionResult(text, execRes.data.readback);
        execRes.data.verificationStatus = verification.status;
        execRes.data.verificationNote = verification.note;
        assistantMsg.execution = execRes.data;

        // 对话气泡最终回复保持简短（1句话），不污染聊天界面
        if (verification.status === 'verified') {
          assistantMsg.content = parsed.explanation
            ? `${parsed.explanation}（区域 ${execRes.data.readback?.usedRangeAddress || ''} 验证通过）`
            : `已在【${targetWbName}】执行完成，区域 ${execRes.data.readback?.usedRangeAddress || ''} 验证通过。`;
        } else {
          assistantMsg.content = parsed.explanation
            ? `${parsed.explanation}（${verification.note}）`
            : `已在【${targetWbName}】执行宏，${verification.note}`;
        }
      } else {
        const errMsg = execRes.error || '执行遇到阻断';
        assistantMsg.content = `执行中断：${errMsg}`;
        assistantMsg.execution = {
          summary: errMsg,
          error: errMsg,
          elapsedMs: 0,
          vbaCode: extracted.code,
        };
      }

      // 刷新工作簿与快照时间轴
      await refreshWorkbookInfo();
    } catch (err: any) {
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

    const targetWbName = workbook?.name || '';
    const targetWbFullName = workbook?.fullName || '';

    const assistantMsgId = 'ai_script_' + Date.now();
    const assistantMsg: ChatMessage = {
      id: assistantMsgId,
      role: 'assistant',
      content: `正在向【${targetWbName || '当前活动工作簿'}】运行本地脚本: ${script.name}...`,
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
        targetWorkbookName: targetWbName,
        targetWorkbookFullName: targetWbFullName,
      });

      if (execRes.ok && execRes.data) {
        const verification = verifyExecutionResult(`运行脚本: ${script.name}`, execRes.data.readback);
        execRes.data.verificationStatus = verification.status;
        execRes.data.verificationNote = verification.note;
        assistantMsg.execution = execRes.data;
        assistantMsg.content = `脚本【${script.name}】已运行完成。${verification.note}`;
      } else {
        assistantMsg.content = `脚本执行失败：${execRes.error}`;
        assistantMsg.execution = {
          summary: execRes.error || '脚本执行失败',
          error: execRes.error,
          elapsedMs: 0,
          vbaCode: script.code,
        };
      }
      await refreshWorkbookInfo();
    } catch (e: any) {
      assistantMsg.content = `运行失败: ${e.message}`;
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
