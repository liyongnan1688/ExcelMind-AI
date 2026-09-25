<script lang="ts">
  import { FileSpreadsheet, Settings, Code2, RefreshCw, History } from 'lucide-svelte';
  import type { WorkbookInfo } from '../services/bridge';

  export let workbook: WorkbookInfo | null = null;
  export let onOpenSettings: () => void;
  export let onOpenScripts: () => void;
  export let onRefresh: () => void;
</script>

<header class="office-header">
  <div class="header-left">
    <div class="app-icon">
      <FileSpreadsheet size={18} color="#107C41" />
    </div>
    <div class="title-container">
      <div class="app-title">Excel AI 助手</div>
      <div class="workbook-subtitle" title={workbook?.fullName || '未检测到工作簿'}>
        {#if workbook}
          <span class="dot"></span>
          <span class="wb-name">{workbook.name}</span>
          {#if !workbook.isSaved}
            <span class="unsaved-tag">未保存</span>
          {/if}
        {:else}
          <span class="no-wb">等待 Excel 连接...</span>
        {/if}
      </div>
    </div>
  </div>

  <div class="header-actions">
    <button class="btn-icon" title="刷新工作簿状态" on:click={onRefresh}>
      <RefreshCw size={15} />
    </button>
    <button class="btn-icon" title="我的脚本库" on:click={onOpenScripts}>
      <Code2 size={16} />
    </button>
    <button class="btn-icon" title="API 设置" on:click={onOpenSettings}>
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
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04);
  }

  .header-left {
    display: flex;
    align-items: center;
    gap: 10px;
    min-width: 0;
  }

  .app-icon {
    width: 28px;
    height: 28px;
    background: var(--excel-light);
    border-radius: 4px;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
  }

  .title-container {
    display: flex;
    flex-direction: column;
    min-width: 0;
  }

  .app-title {
    font-size: 13px;
    font-weight: 600;
    color: var(--office-text);
    line-height: 1.2;
  }

  .workbook-subtitle {
    font-size: 11px;
    color: var(--office-muted);
    display: flex;
    align-items: center;
    gap: 4px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  .wb-name {
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 170px;
  }

  .dot {
    width: 6px;
    height: 6px;
    border-radius: 50%;
    background: var(--excel-green);
    display: inline-block;
  }

  .unsaved-tag {
    font-size: 10px;
    background: #fff4ce;
    color: #795b00;
    padding: 1px 4px;
    border-radius: 2px;
  }

  .header-actions {
    display: flex;
    align-items: center;
    gap: 4px;
  }
</style>
