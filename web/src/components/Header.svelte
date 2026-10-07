<script lang="ts">
  import {
    FileSpreadsheet,
    RefreshCw,
    MoreHorizontal,
    Bot,
    FolderGit2,
    Wrench,
    Layers,
    GitBranch,
    Settings,
    ShieldCheck,
  } from 'lucide-svelte';
  import type { WorkbookInfo } from '../services/bridge';

  export let workbook: WorkbookInfo | null = null;
  export let activeWorkspace: 'assistant' | 'scripts' | 'datatools' = 'assistant';
  export let onSelectWorkspace: (ws: 'assistant' | 'scripts' | 'datatools') => void;
  export let onOpenSettings: () => void;
  export let onOpenBatch: () => void;
  export let onOpenWorkflow: () => void;
  export let onRefresh: () => void;

  let showMoreMenu = false;

  function handleSelect(ws: 'assistant' | 'scripts' | 'datatools') {
    onSelectWorkspace(ws);
    showMoreMenu = false;
  }

  function toggleMoreMenu() {
    showMoreMenu = !showMoreMenu;
  }

  function handleMenuAction(action: () => void) {
    showMoreMenu = false;
    action();
  }
</script>

<svelte:window on:click={(e) => {
  const target = e.target as HTMLElement;
  if (!target.closest('.more-menu-container')) {
    showMoreMenu = false;
  }
}} />

<header class="app-header">
  <!-- 第一行：品牌、工作簿状态与快捷操作 -->
  <div class="header-top-bar">
    <div class="brand-section">
      <div class="brand-logo" title="ExcelMind AI 企业级办公助手">
        <FileSpreadsheet size={16} color="#107c41" />
      </div>
      <span class="brand-name">ExcelMind AI</span>
    </div>

    <!-- 紧凑工作簿胶囊：明确标注为“当前活动工作簿”，绝不与宏锁定的任务目标混淆 -->
    <div class="workbook-capsule" title={`当前活动工作簿：${workbook?.fullName || workbook?.name || '等待 Excel 宿主连接...'}`}>
      <span class="capsule-label">当前:</span>
      {#if workbook && workbook.name && workbook.name !== '未检测到活动工作簿'}
        <span class="status-dot connected" title="已连接到工作簿"></span>
        <span class="wb-filename">{workbook.name}</span>
        {#if !workbook.isSaved}
          <span class="unsaved-badge" title="当前工作簿有未保存更改">未保存</span>
        {/if}
      {:else}
        <span class="status-dot disconnected" title="未检测到连接"></span>
        <span class="wb-filename disconnected-text">未连接</span>
      {/if}
    </div>

    <!-- 顶部辅助工具 -->
    <div class="top-actions">
      <button
        class="header-btn"
        title="刷新当前工作簿与选区状态"
        on:click={onRefresh}
        aria-label="刷新工作簿状态"
      >
        <RefreshCw size={13} />
      </button>

      <!-- 更多功能下拉 -->
      <div class="more-menu-container">
        <button
          class="header-btn {showMoreMenu ? 'active' : ''}"
          title="更多功能与系统管理"
          on:click|stopPropagation={toggleMoreMenu}
          aria-label="更多功能"
        >
          <MoreHorizontal size={15} />
        </button>

        {#if showMoreMenu}
          <div class="dropdown-popover" role="menu">
            <div class="menu-section-label">更多任务</div>
            <button
              class="menu-item"
              role="menuitem"
              on:click={() => handleMenuAction(onOpenBatch)}
            >
              <Layers size={14} class="menu-icon" />
              <span>多工作簿批量处理</span>
            </button>
            <button
              class="menu-item"
              role="menuitem"
              on:click={() => handleMenuAction(onOpenWorkflow)}
            >
              <GitBranch size={14} class="menu-icon" />
              <span>双步骤任务流水线</span>
            </button>

            <div class="menu-divider"></div>

            <div class="menu-section-label">系统管理</div>
            <button
              class="menu-item"
              role="menuitem"
              on:click={() => handleMenuAction(onOpenSettings)}
            >
              <Settings size={14} class="menu-icon" />
              <span>模型服务与参数设置</span>
            </button>
          </div>
        {/if}
      </div>
    </div>
  </div>

  <!-- 第二行：分段工作区导航 (Segmented Workspace Tabs) -->
  <nav class="workspace-nav" aria-label="工作区导航">
    <div class="segmented-bar">
      <button
        class="segmented-tab {activeWorkspace === 'assistant' ? 'active' : ''}"
        on:click={() => handleSelect('assistant')}
      >
        <Bot size={13} />
        <span>AI 助手</span>
      </button>
      <button
        class="segmented-tab {activeWorkspace === 'scripts' ? 'active' : ''}"
        on:click={() => handleSelect('scripts')}
      >
        <FolderGit2 size={13} />
        <span>宏库</span>
      </button>
      <button
        class="segmented-tab {activeWorkspace === 'datatools' ? 'active' : ''}"
        on:click={() => handleSelect('datatools')}
      >
        <Wrench size={13} />
        <span>数据工具</span>
      </button>
    </div>
  </nav>
</header>

<style>
  .app-header {
    background: #ffffff;
    border-bottom: 1px solid var(--office-border);
    display: flex;
    flex-direction: column;
    flex-shrink: 0;
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.03);
    z-index: 20;
    position: relative;
  }

  .header-top-bar {
    height: 38px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 10px;
    gap: 8px;
  }

  .brand-section {
    display: flex;
    align-items: center;
    gap: 6px;
    flex-shrink: 0;
  }

  .brand-logo {
    width: 22px;
    height: 22px;
    border-radius: var(--office-radius-xs);
    background: var(--excel-light);
    border: 1px solid var(--excel-light-border);
    display: flex;
    align-items: center;
    justify-content: center;
  }

  .brand-name {
    font-size: var(--font-size-base);
    font-weight: 600;
    color: var(--office-text);
    letter-spacing: -0.2px;
  }

  .workbook-capsule {
    display: flex;
    align-items: center;
    gap: 5px;
    background: #f8f9fa;
    border: 1px solid var(--office-border-subtle);
    padding: 2px 7px;
    border-radius: var(--office-radius-full);
    min-width: 0;
    max-width: 170px;
    cursor: default;
  }

  .status-dot {
    width: 6px;
    height: 6px;
    border-radius: 50%;
    flex-shrink: 0;
  }

  .status-dot.connected {
    background: var(--excel-green);
    box-shadow: 0 0 0 1.5px rgba(16, 124, 65, 0.2);
  }

  .status-dot.disconnected {
    background: var(--office-dim);
  }

  .capsule-label {
    font-size: 10px;
    color: var(--office-dim);
    font-weight: 600;
    flex-shrink: 0;
    text-transform: uppercase;
  }

  .wb-filename {
    font-size: var(--font-size-xs);
    font-weight: 500;
    color: var(--office-text-secondary);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 95px;
  }

  .disconnected-text {
    color: var(--office-muted);
    font-style: italic;
  }

  .unsaved-badge {
    font-size: 10px;
    background: var(--office-amber-light);
    color: var(--office-amber);
    border: 1px solid var(--office-amber-border);
    padding: 0 3px;
    border-radius: 2px;
    line-height: 1.2;
    flex-shrink: 0;
  }

  .top-actions {
    display: flex;
    align-items: center;
    gap: 3px;
    flex-shrink: 0;
  }

  .header-btn {
    width: 26px;
    height: 26px;
    display: flex;
    align-items: center;
    justify-content: center;
    border: none;
    background: transparent;
    color: var(--office-muted);
    border-radius: var(--office-radius-xs);
    cursor: pointer;
    transition: all 0.15s ease;
  }

  .header-btn:hover {
    background: var(--office-hover);
    color: var(--office-text);
  }

  .header-btn.active {
    background: var(--office-active);
    color: var(--excel-green);
  }

  /* 下拉菜单 */
  .more-menu-container {
    position: relative;
  }

  .dropdown-popover {
    position: absolute;
    top: 30px;
    right: 0;
    width: 180px;
    background: #ffffff;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    box-shadow: 0 6px 16px rgba(0, 0, 0, 0.12);
    padding: 4px;
    z-index: 100;
    display: flex;
    flex-direction: column;
    gap: 2px;
  }

  .menu-section-label {
    font-size: 10px;
    font-weight: 600;
    color: var(--office-dim);
    padding: 4px 8px 2px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }

  .menu-item {
    display: flex;
    align-items: center;
    gap: 8px;
    width: 100%;
    padding: 6px 8px;
    font-size: var(--font-size-sm);
    color: var(--office-text);
    background: transparent;
    border: none;
    border-radius: var(--office-radius-xs);
    cursor: pointer;
    text-align: left;
    transition: background 0.1s ease;
  }

  .menu-item:hover {
    background: var(--office-hover);
    color: var(--excel-green);
  }

  .menu-icon {
    color: var(--office-muted);
    flex-shrink: 0;
  }

  .menu-divider {
    height: 1px;
    background: var(--office-border-subtle);
    margin: 3px 0;
  }

  /* 第二行：分段标签 */
  .workspace-nav {
    padding: 0 8px 6px;
    box-sizing: border-box;
    width: 100%;
  }

  .segmented-bar {
    display: flex;
    align-items: center;
    background: #f0f0f0;
    border-radius: var(--office-radius-sm);
    padding: 2px;
    gap: 2px;
    width: 100%;
    box-sizing: border-box;
  }

  .segmented-tab {
    flex: 1 1 0%;
    min-width: 0;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 3px;
    padding: 5px 2px;
    font-size: var(--font-size-xs);
    font-weight: 500;
    color: var(--office-muted);
    background: transparent;
    border: none;
    border-radius: var(--office-radius-xs);
    cursor: pointer;
    transition: all 0.15s ease;
    box-sizing: border-box;
  }

  .segmented-tab:hover:not(.active) {
    color: var(--office-text);
    background: rgba(0, 0, 0, 0.04);
  }

  .segmented-tab.active {
    background: #ffffff;
    color: var(--excel-green);
    font-weight: 600;
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.08);
  }

  .segmented-tab span {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    font-size: 11px;
  }

  @media (max-width: 340px) {
    .header-top-bar {
      padding: 0 6px;
      gap: 4px;
    }
    .brand-name {
      font-size: 12px;
    }
    .workbook-capsule {
      max-width: 95px;
      padding: 1px 4px;
    }
    .wb-filename {
      max-width: 48px;
    }
    .workspace-nav {
      padding: 0 6px 6px;
    }
    .segmented-tab {
      padding: 4px 1px;
      gap: 2px;
    }
  }
</style>

