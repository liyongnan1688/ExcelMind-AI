<script lang="ts">
  import { FileSpreadsheet, Settings, Code2, RefreshCw } from 'lucide-svelte';
  import type { WorkbookInfo } from '../services/bridge';

  export let workbook: WorkbookInfo | null = null;
  export let onOpenSettings: () => void;
  export let onOpenScripts: () => void;
  export let onRefresh: () => void;
</script>

<header class="office-header">
  <div class="header-left">
    <div class="app-icon" title="Lee-Excel AI 助手">
      <FileSpreadsheet size={18} color="#107C41" />
    </div>
    <div class="title-container">
      <div class="app-title">Excel AI 助手</div>
      <div class="workbook-subtitle" title={workbook?.fullName || '等待 Excel 宿主连接...'}>
        {#if workbook && workbook.name && workbook.name !== '未检测到活动工作簿'}
          <span class="dot" title="已连接到工作簿"></span>
          <span class="wb-name">{workbook.name}</span>
          {#if !workbook.isSaved}
            <span class="unsaved-tag">未保存</span>
          {/if}
        {:else}
          <span class="no-wb">{workbook?.name || '等待 Excel 连接...'}</span>
        {/if}
      </div>
    </div>
  </div>

  <div class="header-actions">
    <button class="btn-icon" title="刷新工作簿状态与快照" on:click={onRefresh} aria-label="刷新">
      <RefreshCw size={15} />
    </button>
    <button class="btn-icon" title="我的脚本库 (.bas)" on:click={onOpenScripts} aria-label="我的脚本库">
      <Code2 size={16} />
    </button>
    <button class="btn-icon" title="大模型 API 设置" on:click={onOpenSettings} aria-label="API 设置">
      <Settings size={16} />
    </button>
  </div>
</header>

<style>
  .office-header {
    height: 48px;
    background: #ffffff;
    border-bottom: 1px solid var(--office-border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 12px;
    box-shadow: var(--office-shadow-sm);
    flex-shrink: 0;
    gap: 8px;
  }

  .header-left {
    display: flex;
    align-items: center;
    gap: 10px;
    min-width: 0;
    flex: 1;
  }

  .app-icon {
    width: 28px;
    height: 28px;
    background: var(--excel-light);
    border-radius: var(--office-radius-sm);
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
    border: 1px solid var(--excel-light-border);
  }

  .title-container {
    display: flex;
    flex-direction: column;
    min-width: 0;
    gap: 1px;
  }

  .app-title {
    font-size: var(--font-size-base);
    font-weight: 600;
    color: var(--office-text);
    line-height: var(--line-height-tight);
    letter-spacing: -0.2px;
  }

  .workbook-subtitle {
    font-size: var(--font-size-xs);
    color: var(--office-muted);
    display: flex;
    align-items: center;
    gap: 5px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    line-height: var(--line-height-tight);
  }

  .wb-name {
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: clamp(130px, 35vw, 260px);
    font-weight: 500;
    color: var(--office-text-secondary);
  }

  .no-wb {
    color: var(--office-dim);
    font-style: italic;
  }

  .dot {
    width: 6px;
    height: 6px;
    border-radius: 50%;
    background: var(--excel-green);
    display: inline-block;
    flex-shrink: 0;
  }

  .unsaved-tag {
    font-size: var(--font-size-xs);
    background: var(--office-amber-light);
    color: var(--office-amber);
    border: 1px solid var(--office-amber-border);
    padding: 0 4px;
    border-radius: var(--office-radius-xs);
    line-height: 1.3;
    font-weight: 500;
    flex-shrink: 0;
  }

  .header-actions {
    display: flex;
    align-items: center;
    gap: 2px;
    flex-shrink: 0;
  }
</style>
