<script lang="ts">
  import { onMount, onDestroy, tick } from 'svelte';
  import { Zap, MessageSquare } from 'lucide-svelte';
  import Header from './components/Header.svelte';
  import SettingsModal from './components/SettingsModal.svelte';
  import ScriptDrawer from './components/ScriptDrawer.svelte';
  import ExecutionCard from './components/ExecutionCard.svelte';
  import ChatInput from './components/ChatInput.svelte';
  import { bridge, type WorkbookInfo, type ScriptItem, type VbaExecutionData, type SelectionSnapshot } from './services/bridge';
  import {
    callLlmStream,
    extractVbaCode,
    parseStreamOutput,
    verifyExecutionResult,
  } from './services/llm';

  interface ChatMessage {
    id: string;
    role: 'user' | 'assistant';
    content: string; // 纯文本内容/简述
    mode?: 'AUTOMATION' | 'CHAT'; // 模式标识
    prompt?: string;
    streamVbaCode?: string; // 流式生成的代码
    execution?: VbaExecutionData | null;
    isExecuting?: boolean;
    selectionSnapshot?: SelectionSnapshot; // 附加的只读选区快照
  }

  let workbook: WorkbookInfo | null = null;
  let messages: ChatMessage[] = [];
  let isProcessing = false;
  let messagesContainer: HTMLElement;
  let inspectingSnapshot: SelectionSnapshot | null = null;

  let showSettings = false;
  let showScripts = false;
  let scriptDrawerRef: any;

  let pollTimer: any = null;
  let unbindWorkbookChange: (() => void) | null = null;
  let unbindOpenScripts: (() => void) | null = null;
  let unbindOpenSettings: (() => void) | null = null;
  let unbindOpenImportMacro: (() => void) | null = null;
  let unbindOpenMyMacros: (() => void) | null = null;
  let unbindOpenMacroLibrary: (() => void) | null = null;

  onMount(() => {
    // 监听 C# 宿主推送的工作簿激活变更
    unbindWorkbookChange = bridge.onWorkbookChange((info) => {
      workbook = info;
    });

    // 监听 Ribbon 菜单打开“宏库”/“我的脚本”抽屉
    unbindOpenMacroLibrary = bridge.onAction('open_macro_library', () => {
      showScripts = true;
      showSettings = false;
      scriptDrawerRef?.openTab('list');
    });

    unbindOpenScripts = bridge.onAction('open_scripts', () => {
      showScripts = true;
      showSettings = false;
      scriptDrawerRef?.openTab('list');
    });

    // 监听 Ribbon 菜单打开“我的宏”列表
    unbindOpenMyMacros = bridge.onAction('open_my_macros', () => {
      showScripts = true;
      showSettings = false;
      scriptDrawerRef?.openTab('list');
    });

    // 监听 Ribbon 菜单打开“导入宏”窗口
    unbindOpenImportMacro = bridge.onAction('open_import_macro', () => {
      showScripts = true;
      showSettings = false;
      scriptDrawerRef?.openTab('import');
    });

    // 监听 Ribbon 菜单打开“API配置”弹窗
    unbindOpenSettings = bridge.onAction('open_settings', () => {
      showSettings = true;
      showScripts = false;
    });

    // 初次获取工作簿信息
    refreshWorkbookInfo();

    // 定时轮询，自动感知工作簿新建/打开/切换
    pollTimer = setInterval(() => {
      refreshWorkbookInfo();
    }, 2500);

    // 默认添加欢迎引导语
    messages = [
      {
        id: 'msg_welcome',
        role: 'assistant',
        content:
          '你好！我是你的 ExcelMind AI 助手。已自动绑定当前目标工作簿。\n你可以用自然语言下达表格操作、数据汇总或进行 Excel 技巧咨询。普通问答直接解答；操作指令在每次运行前自动保存整本物理副本，支持一键回滚。',
      },
    ];

    if (typeof window !== 'undefined') {
      (window as any).__openScripts = () => {
        showScripts = true;
        showSettings = false;
        scriptDrawerRef?.openTab('list');
      };
      (window as any).__openMyMacros = () => {
        showScripts = true;
        showSettings = false;
        scriptDrawerRef?.openTab('list');
      };
      (window as any).__openImportMacro = () => {
        showScripts = true;
        showSettings = false;
        scriptDrawerRef?.openTab('import');
      };
      (window as any).__openSettings = () => {
        showSettings = true;
        showScripts = false;
      };

      (window as any).__addTestExecutionMessage = (promptText: string, executionData: any) => {
        messages = [
          ...messages,
          {
            id: 'msg_test_' + Date.now(),
            role: 'assistant',
            content: executionData?.summary || '自动化执行结果',
            prompt: promptText,
            streamVbaCode: '',
            isExecuting: false,
            execution: executionData,
          },
        ];
      };
    }
  });

  onDestroy(() => {
    if (pollTimer) clearInterval(pollTimer);
    if (unbindWorkbookChange) unbindWorkbookChange();
    if (unbindOpenMacroLibrary) unbindOpenMacroLibrary();
    if (unbindOpenScripts) unbindOpenScripts();
    if (unbindOpenSettings) unbindOpenSettings();
    if (unbindOpenImportMacro) unbindOpenImportMacro();
    if (unbindOpenMyMacros) unbindOpenMyMacros();
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

  // 核心交互：用户显式入口驱动、请求生命周期固定、通道响应分流、目标锁定与写后核验
  async function handleSend(text: string, userSelectedMode?: 'AUTOMATION' | 'CHAT', selectionSnapshot?: SelectionSnapshot) {
    if (!text.trim() || isProcessing) return;

    // 工作簿一致性核验：若附加了选区，且为操作模式，核对目标工作簿
    const requestMode: 'AUTOMATION' | 'CHAT' = userSelectedMode || 'AUTOMATION';
    if (selectionSnapshot && requestMode === 'AUTOMATION') {
      const currentWbName = (workbook?.name || '').trim();
      const attachedWbName = (selectionSnapshot.context.workbookName || '').trim();
      if (currentWbName && attachedWbName && currentWbName.toLowerCase() !== attachedWbName.toLowerCase()) {
        alert(`⚠️ 附加选区来自工作簿【${attachedWbName}】，但当前活动工作簿为【${currentWbName}】。\n\n为避免跨工作簿误操作，请重新附加当前工作簿选区或移除附件后再发送。`);
        return;
      }
    }

    isProcessing = true;

    // 唯一模式来源：用户显式入口，严禁根据文本内容猜测意图
    const requestId = 'req_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7);
    const userMsgId = 'user_' + requestId;
    const assistantMsgId = 'ai_' + requestId;

    // 发送瞬间锁定当前目标工作簿上下文
    const targetWbName = workbook?.name || '';
    const targetWbFullName = workbook?.fullName || '';
    const sheets = workbook?.sheets || [];
    const activeSheet = workbook?.activeSheetName || '';
    const usedRange = workbook?.usedRangeAddress || '';

    const currentRequest = {
      requestId,
      requestMode,
      userMsgId,
      assistantMsgId,
      targetWbName,
      targetWbFullName,
      sheets,
      activeSheet,
      usedRange,
      selectionSnapshot,
    };

    // 1. 添加用户消息（显示发送瞬间锁定的模式标签与选区快照）
    messages = [
      ...messages,
      {
        id: currentRequest.userMsgId,
        role: 'user',
        content: text,
        mode: currentRequest.requestMode,
        selectionSnapshot: currentRequest.selectionSnapshot,
      },
    ];
    await scrollToBottom();

    // 2. 添加助手占位消息
    const assistantMsg: ChatMessage = {
      id: currentRequest.assistantMsgId,
      role: 'assistant',
      content: currentRequest.requestMode === 'CHAT' ? '正在思考解答...' : '正在准备自动化方案...',
      prompt: text,
      streamVbaCode: '',
      isExecuting: true,
      execution: null,
    };
    messages = [...messages, assistantMsg];
    await scrollToBottom();

    try {
      const chatHistory = messages
        .filter((m) => m.id !== currentRequest.assistantMsgId && m.id !== currentRequest.userMsgId && m.id !== 'msg_welcome')
        .map((m) => ({ role: m.role, content: m.content }));

      // 调用大模型流式生成 (严格使用本次请求锁定的 requestMode)
      const streamRes = await callLlmStream(
        text,
        currentRequest.targetWbName,
        currentRequest.sheets,
        currentRequest.activeSheet,
        currentRequest.usedRange,
        currentRequest.requestMode,
        (partialText) => {
          if (currentRequest.requestMode === 'AUTOMATION') {
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
        chatHistory,
        currentRequest.selectionSnapshot
      );

      let rawResponse = streamRes.fullText;
      let finishReason = streamRes.finishReason;
      let currentApiAudit = streamRes.audit;

      // 四、处理 CHAT 通道：完整展示模型回答，绝对不进入自动化执行链
      if (currentRequest.requestMode === 'CHAT') {
        assistantMsg.content = rawResponse;
        assistantMsg.streamVbaCode = '';
        assistantMsg.execution = null;
        assistantMsg.isExecuting = false;
        messages = [...messages];
        return;
      }

      // 五、处理 AUTOMATION 通道：客观区分未取得代码、代码无效与宏实际失败
      let extracted = extractVbaCode(rawResponse, finishReason);

      // 分支 1: 模型返回普通文字，未提供 VBA（如输入“你是”返回了自我介绍）
      if (extracted.status === 'no_code') {
        assistantMsg.content = rawResponse + '\n\n（本次未执行：模型未返回可执行 VBA）';
        assistantMsg.streamVbaCode = '';
        assistantMsg.execution = null; // 绝不创建宏执行失败卡！
        assistantMsg.isExecuting = false;
        messages = [...messages];
        return;
      }

      // 分支 2: 模型返回疑似 VBA，但输出被截断、未闭合或结构缺失（代码未执行）
      if (extracted.status !== 'valid' || !extracted.code) {
        const errorReason = extracted.error || '代码不完整或结构缺失，已安全停止执行';
        const parsed = parseStreamOutput(rawResponse);
        assistantMsg.content = parsed.explanation
          ? `${parsed.explanation}\n\n（代码未执行：${errorReason}）`
          : `（代码未执行：${errorReason}）`;
        assistantMsg.streamVbaCode = extracted.code || '';
        assistantMsg.execution = {
          summary: `代码未执行：${errorReason}`,
          error: errorReason,
          elapsedMs: 0,
          vbaCode: extracted.code || '',
          originalVbaCode: extracted.code || '',
          executedVbaCode: '',
          precheckStatus: extracted.status === 'scope_risk' ? 'scope_risk_intercepted' : 'failed',
          executionPhase: 'intercepted_before_run',
          isSourceIdentical: true,
          apiAudit: currentApiAudit,
        };
        assistantMsg.isExecuting = false;
        messages = [...messages];
        return;
      }

      // 分支 3: 提取成功且结构完整：才进入执行链路 (显式绑定目标工作簿)
      let parsed = parseStreamOutput(rawResponse);
      assistantMsg.streamVbaCode = extracted.code;

      let execRes = await bridge.send<VbaExecutionData>('execute_vba', {
        code: extracted.code,
        prompt: text,
        targetWorkbookName: currentRequest.targetWbName,
        targetWorkbookFullName: currentRequest.targetWbFullName,
        rawModelResponse: rawResponse,
      });

      let wasRepaired = false;

      // 智能重新生成：如果初次代码在宿主编译未通过，大模型结合真实拦截原因重新生成全新代码
      if (!execRes.ok && execRes.data?.precheckStatus === 'failed') {
        assistantMsg.content = '初次代码在 Excel VBE 编译预检未通过，正在请求 AI 分析真实原因并重新生成...';
        messages = [...messages];
        await scrollToBottom();

        const repairPrompt = `此前针对用户指令【${text}】生成的代码在 Excel VBE 静态预编译中未通过，具体诊断信息：【${execRes.data?.error || execRes.error || '存在语法或未定义引用错误'}】。
请结合上述真实报错原因重新生成完整无误、可执行的纯 VBA 源码：
1. 主执行过程建议命名为 Sub Main(targetWb As Workbook) 或 Sub Main()；
2. 确保所有控制结构严格配对闭合（如 For 与 Next、If 与 End If、With 与 End With）；
3. 直接输出纯 VBA 源码，不要输出多余解释或伪代码。`;

        try {
          const repairedStreamRes = await callLlmStream(
            repairPrompt,
            currentRequest.targetWbName,
            currentRequest.sheets,
            currentRequest.activeSheet,
            currentRequest.usedRange,
            'AUTOMATION',
            (partialText) => {
              const p = parseStreamOutput(partialText);
              if (p.hasCode) {
                assistantMsg.content = '正在重新编译并执行新版代码...';
                assistantMsg.streamVbaCode = p.vbaCode;
              }
              messages = [...messages];
            },
            []
          );

          const repairedResponse = repairedStreamRes.fullText;
          const repairedExtracted = extractVbaCode(repairedResponse, repairedStreamRes.finishReason);
          if (!repairedExtracted.error && repairedExtracted.code) {
            rawResponse = repairedResponse;
            extracted = repairedExtracted;
            parsed = parseStreamOutput(repairedResponse);
            assistantMsg.streamVbaCode = repairedExtracted.code;

            const secondExecRes = await bridge.send<VbaExecutionData>('execute_vba', {
              code: repairedExtracted.code,
              prompt: text,
              targetWorkbookName: currentRequest.targetWbName,
              targetWorkbookFullName: currentRequest.targetWbFullName,
              rawModelResponse: repairedResponse,
            });

            if (secondExecRes.ok && secondExecRes.data) {
              execRes = secondExecRes;
              wasRepaired = true;
            } else if (secondExecRes.data) {
              execRes = secondExecRes;
            }
          }
        } catch {
          // 修复网络异常时保留首次执行结果供排查
        }
      }

      if (execRes.ok && execRes.data) {
        // 执行后核验：比对写后读回与用户实际诉求
        const verification = verifyExecutionResult(text, execRes.data.readback);
        execRes.data.verificationStatus = verification.status;
        execRes.data.verificationNote = verification.note;
        execRes.data.apiAudit = currentApiAudit;
        assistantMsg.execution = execRes.data;

        const repairBadge = wasRepaired ? '（初次代码存在预编译瑕疵，已自动完成自我修复并成功执行）' : '';

        // 对话气泡最终回复保持简短（1句话），不污染聊天界面
        if (verification.status === 'verified') {
          assistantMsg.content = parsed.explanation
            ? `${parsed.explanation}（区域 ${execRes.data.readback?.usedRangeAddress || ''} 验证通过）${repairBadge}`
            : `已在【${targetWbName}】执行完成，区域 ${execRes.data.readback?.usedRangeAddress || ''} 验证通过。${repairBadge}`;
        } else {
          assistantMsg.content = parsed.explanation
            ? `${parsed.explanation}（${verification.note}）${repairBadge}`
            : `已在【${targetWbName}】执行宏，${verification.note}${repairBadge}`;
        }
      } else {
        const errMsg = execRes.error || '执行遇到阻断';
        assistantMsg.content = `执行中断：${errMsg}`;
        assistantMsg.execution = execRes.data
          ? {
              ...execRes.data,
              apiAudit: currentApiAudit,
              summary: errMsg,
              error: errMsg,
              vbaCode: extracted.code,
            }
          : {
              summary: errMsg,
              error: errMsg,
              elapsedMs: 0,
              vbaCode: extracted.code,
              apiAudit: currentApiAudit,
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

  // 从“我的宏”直接运行
  async function handleRunScript(script: ScriptItem, entryPoint?: string) {
    if (isProcessing) return;
    isProcessing = true;

    const chosenEntryPoint = entryPoint || script.entryPoint || '';
    const macroDisplayName = script.displayName || script.name;
    const entryLabel = chosenEntryPoint ? ` (入口: ${chosenEntryPoint})` : '';

    const userMsgId = 'user_script_' + Date.now();
    messages = [
      ...messages,
      {
        id: userMsgId,
        role: 'user',
        content: `运行宏: 【${macroDisplayName}】${entryLabel}`,
      },
    ];

    const targetWbName = workbook?.name || '';
    const targetWbFullName = workbook?.fullName || '';

    const assistantMsgId = 'ai_script_' + Date.now();
    const assistantMsg: ChatMessage = {
      id: assistantMsgId,
      role: 'assistant',
      content: `正在向【${targetWbName || '当前活动工作簿'}】运行宏: ${macroDisplayName}...`,
      prompt: `运行宏: ${macroDisplayName}`,
      isExecuting: true,
      execution: null,
    };
    messages = [...messages, assistantMsg];
    await scrollToBottom();

    try {
      const execRes = await bridge.send<VbaExecutionData>('execute_vba', {
        code: script.code,
        prompt: `运行宏: ${macroDisplayName}`,
        targetWorkbookName: targetWbName,
        targetWorkbookFullName: targetWbFullName,
        entryPoint: chosenEntryPoint,
        scriptId: script.id || script.fileName,
      });

      if (execRes.ok && execRes.data) {
        const verification = verifyExecutionResult(`运行宏: ${macroDisplayName}`, execRes.data.readback);
        execRes.data.verificationStatus = verification.status;
        execRes.data.verificationNote = verification.note;
        assistantMsg.execution = execRes.data;
        assistantMsg.content = `宏【${macroDisplayName}】已运行完成。${verification.note}`;
      } else {
        assistantMsg.content = `宏执行失败：${execRes.error}`;
        assistantMsg.execution = {
          summary: execRes.error || '宏执行失败',
          error: execRes.error,
          elapsedMs: 0,
          vbaCode: script.code,
        };
      }
      await refreshWorkbookInfo();
      scriptDrawerRef?.refreshScripts();
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
          <div class="user-bubble-wrapper">
            {#if msg.mode}
              <div class="user-mode-tag {msg.mode === 'AUTOMATION' ? 'tag-auto' : 'tag-chat'}">
                {#if msg.mode === 'AUTOMATION'}
                  <Zap size={10} />
                  <span>操作指令</span>
                {:else}
                  <MessageSquare size={10} />
                  <span>咨询对话</span>
                {/if}
              </div>
            {/if}
            <div class="user-bubble {msg.mode === 'CHAT' ? 'bubble-chat' : 'bubble-auto'}">{msg.content}</div>

            <!-- 若随请求附加了选区，在用户气泡下方显示紧凑标识与查看发送按钮 -->
            {#if msg.selectionSnapshot}
              <div class="user-selection-badge">
                <span class="badge-icon">📎</span>
                <span class="badge-text">
                  已附加选区: {msg.selectionSnapshot.context.sheetName}!{msg.selectionSnapshot.context.address}
                  ({msg.selectionSnapshot.options.includeSamples ? `含${msg.selectionSnapshot.context.sampleRowCount}行样本` : '仅结构'})
                </span>
                <button
                  class="badge-view-btn"
                  type="button"
                  on:click={() => (inspectingSnapshot = msg.selectionSnapshot)}
                  title="查看本次随提示词发送给模型的只读数据"
                >
                  查看发送内容
                </button>
              </div>
            {/if}
          </div>
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
              on:restored={refreshWorkbookInfo}
              on:expand={scrollToBottom}
            />
          {/if}
        </div>
      {/if}
    {/each}
  </main>

  <!-- 底部紧凑输入栏 -->
  <ChatInput disabled={isProcessing} onSend={handleSend} />

  <!-- 查看发送内容弹窗 (脱敏展示，绝不暴露 API Key) -->
  {#if inspectingSnapshot}
    <div class="modal-backdrop" role="presentation" on:click|self={() => (inspectingSnapshot = null)}>
      <div class="modal-card">
        <div class="modal-header">
          <h3>📎 本次随请求发送的只读选区数据</h3>
          <button class="close-x-btn" on:click={() => (inspectingSnapshot = null)}>✕</button>
        </div>
        <div class="modal-body">
          <div class="audit-summary-bar">
            <span>包含字段: <code>{inspectingSnapshot.summary.fieldsIncluded.join(', ')}</code></span>
            <span>样本范围: <code>{inspectingSnapshot.summary.sampleRowRange}</code></span>
            <span>字符数: <code>{inspectingSnapshot.summary.totalChars}</code></span>
          </div>
          <p class="modal-tip">以下为作为低信任数据段注入请求的完整文本（不含任何系统 API Key 凭据）：</p>
          <pre class="payload-preview-code"><code>{inspectingSnapshot.formattedText}</code></pre>
        </div>
        <div class="modal-footer">
          <button class="primary-btn" on:click={() => (inspectingSnapshot = null)}>关闭</button>
        </div>
      </div>
    </div>
  {/if}

  <!-- 设置弹窗 -->
  {#if showSettings}
    <SettingsModal onClose={() => (showSettings = false)} />
  {/if}

  <!-- “宏管理与我的脚本”抽屉 -->
  <ScriptDrawer
    bind:this={scriptDrawerRef}
    isOpen={showScripts}
    {workbook}
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
    scroll-behavior: smooth;
  }

  .message-row {
    display: flex;
    flex-direction: column;
    max-width: 100%;
  }

  .user-row {
    align-items: flex-end;
  }

  .user-bubble-wrapper {
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    max-width: 86%;
    gap: 4px;
  }

  .user-mode-tag {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    font-size: var(--font-size-xs);
    font-weight: 500;
    padding: 1px 7px;
    border-radius: var(--office-radius-full);
    line-height: var(--line-height-tight);
  }

  .user-mode-tag.tag-auto {
    background: var(--excel-light);
    color: var(--excel-dark);
    border: 1px solid var(--excel-light-border);
  }

  .user-mode-tag.tag-chat {
    background: var(--office-blue-light);
    color: var(--office-blue-dark);
    border: 1px solid var(--office-blue-border);
  }

  .user-bubble {
    color: #ffffff;
    padding: 8px 12px;
    border-radius: var(--office-radius) var(--office-radius) 2px var(--office-radius);
    font-size: var(--font-size-base);
    font-family: var(--font-family-ui);
    line-height: var(--line-height-normal);
    width: fit-content;
    word-break: break-all;
    box-shadow: var(--office-shadow-sm);
  }

  .user-bubble.bubble-auto {
    background: var(--excel-green);
  }

  .user-bubble.bubble-chat {
    background: var(--office-blue);
  }

  .ai-row {
    align-items: flex-start;
  }

  .ai-text {
    background: #ffffff;
    padding: 9px 13px;
    border-radius: var(--office-radius) var(--office-radius) var(--office-radius) 2px;
    font-size: var(--font-size-base);
    font-family: var(--font-family-ui);
    line-height: var(--line-height-normal);
    color: var(--office-text);
    border: 1px solid var(--office-border);
    max-width: 95%;
    white-space: pre-wrap;
    word-break: break-all;
    box-shadow: var(--office-shadow-sm);
  }

  /* 附加选区用户气泡徽章 */
  .user-selection-badge {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    background: #f3f9f5;
    border: 1px solid #c7e0d2;
    padding: 2px 6px;
    border-radius: 4px;
    font-size: 11px;
    color: #107c41;
  }

  .badge-icon {
    font-size: 11px;
  }

  .badge-text {
    font-family: var(--font-family-mono, monospace);
  }

  .badge-view-btn {
    border: none;
    background: transparent;
    color: #0078d4;
    cursor: pointer;
    font-size: 11px;
    text-decoration: underline;
    padding: 0 2px;
  }

  .badge-view-btn:hover {
    color: #106ebe;
  }

  /* 模态弹窗 */
  .modal-backdrop {
    position: fixed;
    top: 0;
    left: 0;
    width: 100vw;
    height: 100vh;
    background: rgba(0, 0, 0, 0.4);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 1000;
  }

  .modal-card {
    background: white;
    border-radius: var(--office-radius);
    box-shadow: var(--office-shadow-lg, 0 8px 16px rgba(0, 0, 0, 0.14));
    width: 90%;
    max-width: 400px;
    max-height: 80vh;
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }

  .modal-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 10px 14px;
    border-bottom: 1px solid var(--office-border);
  }

  .modal-header h3 {
    margin: 0;
    font-size: 13px;
    color: var(--office-text);
  }

  .close-x-btn {
    border: none;
    background: transparent;
    font-size: 14px;
    cursor: pointer;
    color: #605e5c;
  }

  .modal-body {
    padding: 12px 14px;
    overflow-y: auto;
    display: flex;
    flex-direction: column;
    gap: 8px;
  }

  .audit-summary-bar {
    display: flex;
    flex-direction: column;
    gap: 4px;
    font-size: 11px;
    background: #f3f2f1;
    padding: 6px 8px;
    border-radius: 4px;
  }

  .audit-summary-bar code {
    background: #e1dfdd;
    padding: 1px 3px;
    border-radius: 2px;
  }

  .modal-tip {
    font-size: 11px;
    color: #605e5c;
    margin: 0;
  }

  .payload-preview-code {
    background: #f8f9fa;
    border: 1px solid #e1dfdd;
    padding: 8px;
    border-radius: 4px;
    font-size: 10px;
    font-family: var(--font-family-mono, monospace);
    white-space: pre-wrap;
    word-break: break-all;
    max-height: 220px;
    overflow-y: auto;
    margin: 0;
  }

  .modal-footer {
    padding: 8px 14px;
    border-top: 1px solid var(--office-border);
    display: flex;
    justify-content: flex-end;
  }

  .primary-btn {
    background: var(--excel-green);
    color: white;
    border: none;
    border-radius: var(--office-radius-sm);
    padding: 4px 14px;
    font-size: 12px;
    cursor: pointer;
  }
</style>
