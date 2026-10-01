<script lang="ts">
  import { Sparkles, Zap, MessageSquare } from 'lucide-svelte';

  export let disabled = false;
  export let onSend: (text: string, mode?: 'AUTOMATION' | 'CHAT') => void;

  let inputText = '';
  let currentMode: 'AUTOMATION' | 'CHAT' = 'AUTOMATION';

  const AUTO_PROMPTS = [
    '用_分裂D列的名称，不要覆盖后面的列，新增。',
    '为首行表头添加浅绿底色并加粗居中',
    '在数据末尾添加汇总行并计算求和公式',
    '将选区所有负数单元格填充为浅红标注',
    '自动调整所有列宽以适应文字长度',
  ];

  const CHAT_PROMPTS = [
    'TEXTSPLIT函数怎么拆分文本？',
    'VLOOKUP与XLOOKUP有哪些区别？',
    '分列时如何避免覆盖右侧已有数据？',
    '如何用INDEX+MATCH实现多条件查找？',
    '解释数据透视表的计算原理与技巧',
  ];

  $: activePrompts = currentMode === 'AUTOMATION' ? AUTO_PROMPTS : CHAT_PROMPTS;

  function handleSubmit(overrideMode?: 'AUTOMATION' | 'CHAT') {
    if (!inputText.trim() || disabled) return;
    const text = inputText.trim();
    const modeToSend = overrideMode || currentMode;
    if (overrideMode && overrideMode !== currentMode) {
      currentMode = overrideMode;
    }
    inputText = '';
    onSend(text, modeToSend);
  }

  function handleKeyDown(e: KeyboardEvent) {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSubmit();
    }
  }

  function setMode(mode: 'AUTOMATION' | 'CHAT') {
    currentMode = mode;
  }

  function handleSelectQuickPrompt(p: string) {
    if (disabled) return;
    inputText = p;
  }
</script>

<div class="input-container">
  <!-- 模式切换与状态栏 -->
  <div class="mode-header">
    <div class="mode-segmented" role="tablist" aria-label="对话模式选择">
      <button
        class="mode-tab {currentMode === 'AUTOMATION' ? 'active-auto' : ''}"
        on:click={() => setMode('AUTOMATION')}
        disabled={disabled}
        type="button"
        role="tab"
        aria-selected={currentMode === 'AUTOMATION'}
      >
        <Zap size={13} class="mode-icon" />
        <span class="mode-title">操作</span>
      </button>
      <button
        class="mode-tab {currentMode === 'CHAT' ? 'active-chat' : ''}"
        on:click={() => setMode('CHAT')}
        disabled={disabled}
        type="button"
        role="tab"
        aria-selected={currentMode === 'CHAT'}
      >
        <MessageSquare size={13} class="mode-icon" />
        <span class="mode-title">对话</span>
      </button>
    </div>

    <div class="mode-tip-text">
      {#if currentMode === 'AUTOMATION'}
        <span class="tip-badge auto-badge">⚡ 操作模式：生成代码并执行，直接修改表格</span>
      {:else}
        <span class="tip-badge chat-badge">💬 对话模式：纯文本咨询指导，不改动表格</span>
      {/if}
    </div>
  </div>

  <!-- 快捷提示气泡 -->
  <div class="quick-prompts">
    <div class="prompts-scroll">
      {#each activePrompts as p}
        <button
          class="prompt-chip {currentMode === 'AUTOMATION' ? 'chip-auto' : 'chip-chat'}"
          on:click={() => handleSelectQuickPrompt(p)}
          disabled={disabled}
          type="button"
        >
          <Sparkles size={11} color={currentMode === 'AUTOMATION' ? '#107C41' : '#0078D4'} />
          <span>{p}</span>
        </button>
      {/each}
    </div>
  </div>

  <!-- 输入主框 -->
  <div class="input-box {currentMode === 'AUTOMATION' ? 'focus-auto' : 'focus-chat'}">
    <textarea
      placeholder={currentMode === 'AUTOMATION'
        ? "输入操作指令 (如: 用_分裂D列的名称，不要覆盖后面的列，新增)..."
        : "向 AI 咨询 Excel 问题或技巧 (如: TEXTSPLIT与分列区别、公式怎么写)..."}
      bind:value={inputText}
      on:keydown={handleKeyDown}
      disabled={disabled}
      rows="2"
    ></textarea>

    <div class="input-actions">
      <div class="shortcut-tip">
        <span class="key-pill">Enter</span> 发送 ({currentMode === 'AUTOMATION' ? '操作' : '对话'})
        <span class="key-sep">/</span>
        <span class="key-pill">Shift+Enter</span> 换行
      </div>
      <div class="action-buttons-group">
        <!-- 对话按钮 -->
        <button
          class="mode-action-btn {currentMode === 'CHAT' ? 'primary-chat' : 'secondary-chat'}"
          on:click={() => handleSubmit('CHAT')}
          disabled={disabled || !inputText.trim()}
          title="以【对话】发送：仅解答，不改动表格"
          type="button"
        >
          <MessageSquare size={13} />
          <span>对话</span>
        </button>

        <!-- 操作按钮 -->
        <button
          class="mode-action-btn {currentMode === 'AUTOMATION' ? 'primary-auto' : 'secondary-auto'}"
          on:click={() => handleSubmit('AUTOMATION')}
          disabled={disabled || !inputText.trim()}
          title="以【操作】发送：直接修改表格并自动备份"
          type="button"
        >
          <Zap size={13} />
          <span>操作</span>
        </button>
      </div>
    </div>
  </div>
</div>

<style>
  .input-container {
    background: #ffffff;
    border-top: 1px solid var(--office-border);
    padding: 8px 12px;
    display: flex;
    flex-direction: column;
    gap: 8px;
    box-shadow: 0 -1px 3px rgba(0, 0, 0, 0.03);
  }

  /* 顶部模式切换栏 */
  .mode-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
  }

  .mode-segmented {
    display: inline-flex;
    background: #f0f2f5;
    border: 1px solid #e1dfdd;
    border-radius: 6px;
    padding: 2px;
    gap: 2px;
  }

  .mode-tab {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 500;
    border: none;
    border-radius: 4px;
    background: transparent;
    color: var(--office-muted);
    cursor: pointer;
    transition: all 0.15s ease;
  }

  .mode-tab:hover:not(:disabled) {
    color: var(--office-text);
  }

  .mode-tab.active-auto {
    background: var(--excel-green);
    color: #ffffff;
    font-weight: 600;
    box-shadow: 0 1px 3px rgba(16, 124, 65, 0.3);
  }

  .mode-tab.active-chat {
    background: #0078d4;
    color: #ffffff;
    font-weight: 600;
    box-shadow: 0 1px 3px rgba(0, 120, 212, 0.3);
  }

  .mode-tab:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }

  .mode-tip-text {
    font-size: 11px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .tip-badge {
    display: inline-flex;
    align-items: center;
    padding: 2px 6px;
    border-radius: 4px;
    font-size: 10.5px;
  }

  .tip-badge.auto-badge {
    background: #e7f3ec;
    color: #0b5a2f;
    border: 1px solid #c2e2cc;
  }

  .tip-badge.chat-badge {
    background: #e8f3fb;
    color: #004e8c;
    border: 1px solid #c7e0f4;
  }

  /* 快捷提示 */
  .quick-prompts {
    overflow-x: auto;
    white-space: nowrap;
    scrollbar-width: none;
  }

  .quick-prompts::-webkit-scrollbar {
    display: none;
  }

  .prompts-scroll {
    display: inline-flex;
    gap: 6px;
  }

  .prompt-chip {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    padding: 3px 8px;
    background: #f8f8f8;
    border: 1px solid var(--office-border);
    border-radius: 12px;
    font-size: 11px;
    color: var(--office-text);
    cursor: pointer;
    transition: all 0.15s;
  }

  .prompt-chip.chip-auto:hover:not(:disabled) {
    background: var(--excel-light);
    border-color: #c2e2cc;
    color: var(--excel-green);
  }

  .prompt-chip.chip-chat:hover:not(:disabled) {
    background: #e8f3fb;
    border-color: #c7e0f4;
    color: #0078d4;
  }

  .prompt-chip:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }

  /* 输入框 */
  .input-box {
    border: 1px solid var(--office-border);
    border-radius: 6px;
    background: white;
    display: flex;
    flex-direction: column;
    transition: all 0.15s ease;
  }

  .input-box.focus-auto:focus-within {
    border-color: var(--excel-green);
    box-shadow: 0 0 0 1px var(--excel-green);
  }

  .input-box.focus-chat:focus-within {
    border-color: #0078d4;
    box-shadow: 0 0 0 1px #0078d4;
  }

  textarea {
    width: 100%;
    resize: none;
    border: none;
    outline: none;
    padding: 8px;
    font-size: 12px;
    font-family: inherit;
    line-height: 1.45;
    color: var(--office-text);
  }

  textarea:disabled {
    background: #faf9f8;
  }

  /* 底部操作行 */
  .input-actions {
    padding: 5px 8px 5px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #faf9f8;
    border-top: 1px solid #f3f2f1;
    border-bottom-left-radius: 6px;
    border-bottom-right-radius: 6px;
    gap: 8px;
  }

  .shortcut-tip {
    font-size: 10.5px;
    color: var(--office-muted);
    display: flex;
    align-items: center;
    gap: 3px;
  }

  .key-pill {
    display: inline-block;
    padding: 1px 4px;
    background: #edebe9;
    border-radius: 3px;
    font-size: 9.5px;
    font-weight: 500;
    color: #323130;
  }

  .key-sep {
    color: #c8c6c4;
    margin: 0 2px;
  }

  .action-buttons-group {
    display: inline-flex;
    align-items: center;
    gap: 6px;
  }

  .mode-action-btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 4px;
    padding: 4px 10px;
    font-size: 11.5px;
    font-weight: 500;
    border-radius: 4px;
    cursor: pointer;
    transition: all 0.15s ease;
  }

  .mode-action-btn:disabled {
    opacity: 0.5;
    cursor: not-allowed;
  }

  /* 操作模式主按钮 */
  .primary-auto {
    background: var(--excel-green);
    border: 1px solid var(--excel-green);
    color: #ffffff;
    box-shadow: 0 1px 2px rgba(16, 124, 65, 0.2);
  }

  .primary-auto:hover:not(:disabled) {
    background: var(--excel-dark);
    border-color: var(--excel-dark);
  }

  /* 操作模式次按钮 (当前是对话时) */
  .secondary-auto {
    background: #ffffff;
    border: 1px solid #c2e2cc;
    color: var(--excel-green);
  }

  .secondary-auto:hover:not(:disabled) {
    background: var(--excel-light);
  }

  /* 对话模式主按钮 */
  .primary-chat {
    background: #0078d4;
    border: 1px solid #0078d4;
    color: #ffffff;
    box-shadow: 0 1px 2px rgba(0, 120, 212, 0.2);
  }

  .primary-chat:hover:not(:disabled) {
    background: #005a9e;
    border-color: #005a9e;
  }

  /* 对话模式次按钮 (当前是操作时) */
  .secondary-chat {
    background: #ffffff;
    border: 1px solid #c7e0f4;
    color: #0078d4;
  }

  .secondary-chat:hover:not(:disabled) {
    background: #e8f3fb;
  }
</style>
