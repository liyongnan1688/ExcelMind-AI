<script lang="ts">
  import { Send, Sparkles } from 'lucide-svelte';

  export let disabled = false;
  export let onSend: (text: string) => void;

  let inputText = '';

  const QUICK_PROMPTS = [
    '为首行表头添加浅绿底色并加粗居中',
    '在数据末尾添加汇总行并计算求和公式',
    '将选区所有负数单元格填充为浅红标注',
    '自动调整所有列宽以适应文字长度',
  ];

  function handleSubmit() {
    if (!inputText.trim() || disabled) return;
    const text = inputText.trim();
    inputText = '';
    onSend(text);
  }

  function handleKeyDown(e: KeyboardEvent) {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSubmit();
    }
  }

  function handleSelectQuickPrompt(p: string) {
    if (disabled) return;
    inputText = p;
  }
</script>

<div class="input-container">
  <!-- 快捷提示气泡 -->
  <div class="quick-prompts">
    <div class="prompts-scroll">
      {#each QUICK_PROMPTS as p}
        <button class="prompt-chip" on:click={() => handleSelectQuickPrompt(p)} disabled={disabled}>
          <Sparkles size={11} color="#107C41" />
          <span>{p}</span>
        </button>
      {/each}
    </div>
  </div>

  <!-- 输入主框 -->
  <div class="input-box">
    <textarea
      placeholder="用自然语言下达 Excel 操作要求 (如: 汇总A列销售额)..."
      bind:value={inputText}
      on:keydown={handleKeyDown}
      disabled={disabled}
      rows="2"
    ></textarea>

    <div class="input-actions">
      <span class="shortcut-tip">Enter 发送 / Shift+Enter 换行</span>
      <button class="btn btn-primary send-btn" on:click={handleSubmit} disabled={disabled || !inputText.trim()}>
        <Send size={13} />
        <span>发送</span>
      </button>
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

  .prompt-chip:hover:not(:disabled) {
    background: var(--excel-light);
    border-color: #c2e2cc;
    color: var(--excel-green);
  }

  .prompt-chip:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }

  .input-box {
    border: 1px solid var(--office-border);
    border-radius: 4px;
    background: white;
    display: flex;
    flex-direction: column;
    transition: border-color 0.15s;
  }

  .input-box:focus-within {
    border-color: var(--excel-green);
    box-shadow: 0 0 0 1px var(--excel-green);
  }

  textarea {
    width: 100%;
    resize: none;
    border: none;
    outline: none;
    padding: 8px;
    font-size: 12px;
    font-family: inherit;
    line-height: 1.4;
    color: var(--office-text);
  }

  textarea:disabled {
    background: #faf9f8;
  }

  .input-actions {
    padding: 4px 8px 6px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #faf9f8;
    border-top: 1px solid #f3f2f1;
    border-bottom-left-radius: 4px;
    border-bottom-right-radius: 4px;
  }

  .shortcut-tip {
    font-size: 10px;
    color: var(--office-muted);
  }

  .send-btn {
    padding: 3px 10px;
    font-size: 11px;
  }
</style>
