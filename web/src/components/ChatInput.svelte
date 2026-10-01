<script lang="ts">
  import { Zap, MessageSquare } from 'lucide-svelte';

  export let disabled = false;
  export let onSend: (text: string, mode?: 'AUTOMATION' | 'CHAT') => void;

  let inputText = '';
  let currentMode: 'AUTOMATION' | 'CHAT' = 'AUTOMATION';

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
        title="操作模式：生成代码并执行，直接修改表格"
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
        title="对话模式：纯文本咨询指导，不改动表格"
      >
        <MessageSquare size={13} class="mode-icon" />
        <span class="mode-title">对话</span>
      </button>
    </div>

    <div class="mode-tip-text">
      {#if currentMode === 'AUTOMATION'}
        <span class="tip-badge auto-badge">⚡ 操作模式：生成代码直接修改表格</span>
      {:else}
        <span class="tip-badge chat-badge">💬 对话模式：纯文本咨询，不改动表格</span>
      {/if}
    </div>
  </div>

  <!-- 输入主框 (加高输入区域，提供更充足的书写空间) -->
  <div class="input-box {currentMode === 'AUTOMATION' ? 'focus-auto' : 'focus-chat'}">
    <textarea
      placeholder={currentMode === 'AUTOMATION'
        ? "输入操作指令 (如: 用_分裂D列的名称，不要覆盖后面的列，新增)..."
        : "向 AI 咨询 Excel 问题或技巧 (如: TEXTSPLIT与分列区别、公式怎么写)..."}
      bind:value={inputText}
      on:keydown={handleKeyDown}
      disabled={disabled}
      rows="3"
    ></textarea>

    <div class="input-actions">
      <div class="shortcut-tip">
        <span class="key-pill">Enter</span> 发送
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
    padding: 10px 12px;
    display: flex;
    flex-direction: column;
    gap: 8px;
    box-shadow: 0 -1px 3px rgba(0, 0, 0, 0.03);
    flex-shrink: 0;
  }

  /* 顶部模式切换栏 */
  .mode-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    min-height: 26px;
  }

  .mode-segmented {
    display: inline-flex;
    background: #f0f2f5;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    padding: 2px;
    gap: 2px;
    flex-shrink: 0;
  }

  .mode-tab {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    padding: 3px 10px;
    font-size: var(--font-size-xs);
    font-weight: 500;
    border: none;
    border-radius: var(--office-radius-sm);
    background: transparent;
    color: var(--office-muted);
    cursor: pointer;
    transition: all 0.15s ease;
    line-height: var(--line-height-tight);
  }

  .mode-tab:hover:not(:disabled) {
    color: var(--office-text);
  }

  .mode-tab.active-auto {
    background: var(--excel-green);
    color: #ffffff;
    font-weight: 600;
    box-shadow: 0 1px 3px rgba(16, 124, 65, 0.25);
  }

  .mode-tab.active-chat {
    background: var(--office-blue);
    color: #ffffff;
    font-weight: 600;
    box-shadow: 0 1px 3px rgba(0, 120, 212, 0.25);
  }

  .mode-tab:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }

  .mode-tip-text {
    font-size: var(--font-size-xs);
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    min-width: 0;
  }

  .tip-badge {
    display: inline-flex;
    align-items: center;
    padding: 2px 7px;
    border-radius: var(--office-radius-xs);
    font-size: var(--font-size-xs);
    line-height: var(--line-height-tight);
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .tip-badge.auto-badge {
    background: var(--excel-light);
    color: var(--excel-dark);
    border: 1px solid var(--excel-light-border);
  }

  .tip-badge.chat-badge {
    background: var(--office-blue-light);
    color: var(--office-blue-dark);
    border: 1px solid var(--office-blue-border);
  }

  /* 输入框 (加高设计与舒适边距) */
  .input-box {
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    background: #ffffff;
    display: flex;
    flex-direction: column;
    transition: all 0.15s ease;
  }

  .input-box.focus-auto:focus-within {
    border-color: var(--excel-green);
    box-shadow: 0 0 0 2px rgba(16, 124, 65, 0.15);
  }

  .input-box.focus-chat:focus-within {
    border-color: var(--office-blue);
    box-shadow: 0 0 0 2px rgba(0, 120, 212, 0.15);
  }

  textarea {
    width: 100%;
    resize: none;
    border: none;
    outline: none;
    padding: 10px 12px;
    font-size: var(--font-size-base);
    font-family: var(--font-family-ui);
    line-height: var(--line-height-normal);
    color: var(--office-text);
    min-height: 82px;
    max-height: 180px;
    background: transparent;
    overflow-y: auto;
  }

  textarea:disabled {
    background: #faf9f8;
    color: var(--office-muted);
  }

  /* 底部操作行 */
  .input-actions {
    padding: 6px 10px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #faf9f8;
    border-top: 1px solid var(--office-border-subtle);
    border-bottom-left-radius: calc(var(--office-radius) - 1px);
    border-bottom-right-radius: calc(var(--office-radius) - 1px);
    gap: 8px;
  }

  .shortcut-tip {
    font-size: var(--font-size-xs);
    color: var(--office-muted);
    display: flex;
    align-items: center;
    gap: 3px;
    white-space: nowrap;
    overflow: hidden;
  }

  .key-pill {
    display: inline-block;
    padding: 1px 4px;
    background: #edebe9;
    border-radius: var(--office-radius-xs);
    font-size: 10px;
    font-weight: 500;
    color: #323130;
    line-height: 1.2;
    border: 1px solid #e1dfdd;
  }

  .key-sep {
    color: #c8c6c4;
    margin: 0 1px;
  }

  .action-buttons-group {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    flex-shrink: 0;
  }

  .mode-action-btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 4px;
    padding: 4px 12px;
    font-size: var(--font-size-sm);
    font-weight: 500;
    border-radius: var(--office-radius-sm);
    cursor: pointer;
    transition: all 0.15s ease;
    line-height: 1.3;
    white-space: nowrap;
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
    background: var(--excel-hover-bg);
    border-color: var(--excel-hover-bg);
  }

  /* 操作模式次按钮 (当前是对话时) */
  .secondary-auto {
    background: #ffffff;
    border: 1px solid var(--excel-light-border);
    color: var(--excel-green);
  }

  .secondary-auto:hover:not(:disabled) {
    background: var(--excel-light);
  }

  /* 对话模式主按钮 */
  .primary-chat {
    background: var(--office-blue);
    border: 1px solid var(--office-blue);
    color: #ffffff;
    box-shadow: 0 1px 2px rgba(0, 120, 212, 0.2);
  }

  .primary-chat:hover:not(:disabled) {
    background: var(--office-blue-dark);
    border-color: var(--office-blue-dark);
  }

  /* 对话模式次按钮 (当前是操作时) */
  .secondary-chat {
    background: #ffffff;
    border: 1px solid var(--office-blue-border);
    color: var(--office-blue);
  }

  .secondary-chat:hover:not(:disabled) {
    background: var(--office-blue-light);
  }
</style>
