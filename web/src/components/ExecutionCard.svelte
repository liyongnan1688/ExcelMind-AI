<script lang="ts">
  import {
    CheckCircle2,
    XCircle,
    Loader2,
    ChevronDown,
    ChevronUp,
    Copy,
    Save,
    RotateCcw,
    Code,
    FileText,
    History,
    Check,
  } from 'lucide-svelte';
  import { bridge, type SnapshotItem, type VbaExecutionData } from '../services/bridge';

  export let prompt: string;
  export let execution: VbaExecutionData | null = null;
  export let streamCode: string = '';
  export let isExecuting = false;
  export let allSnapshots: SnapshotItem[] = [];
  export let onSaveScriptSuccess: () => void;

  // 严格默认折叠，绝不主动铺开，保证窗口极致简洁
  let isExpanded = false;
  let activeTab: 'code' | 'log' = 'code';

  $: displayCode = execution?.vbaCode || streamCode || '';

  let showSaveDialog = false;
  let scriptName = '';
  let scriptDesc = '';
  let isSaving = false;
  let saveSuccess = false;

  let isRestoring = false;
  let restoreSuccess = false;
  let selectedSnapshotId = execution?.snapshot?.id || (allSnapshots[0]?.id ?? '');

  $: if (execution?.snapshot?.id) {
    selectedSnapshotId = execution.snapshot.id;
  }

  let copySuccess = false;
  function handleCopy() {
    if (!displayCode) return;
    navigator.clipboard.writeText(displayCode);
    copySuccess = true;
    setTimeout(() => {
      copySuccess = false;
    }, 1500);
  }

  async function handleSaveScript() {
    if (!scriptName.trim() || !displayCode) return;
    isSaving = true;
    const res = await bridge.send('save_script', {
      name: scriptName.trim(),
      code: displayCode,
      description: scriptDesc.trim() || prompt,
    });
    isSaving = false;

    if (res.ok) {
      saveSuccess = true;
      onSaveScriptSuccess?.();
      setTimeout(() => {
        saveSuccess = false;
        showSaveDialog = false;
        scriptName = '';
      }, 1000);
    } else {
      alert('保存失败: ' + res.error);
    }
  }

  async function handleRestore(targetId: string) {
    if (!targetId) return;
    if (!confirm('确认将当前工作簿恢复到该快照执行前的状态吗？\n当前工作簿未保存的修改将被放弃。')) {
      return;
    }

    isRestoring = true;
    const res = await bridge.send('restore_snapshot', { snapshotId: targetId });
    isRestoring = false;

    if (res.ok) {
      restoreSuccess = true;
      setTimeout(() => {
        restoreSuccess = false;
      }, 2500);
    } else {
      alert('恢复失败: ' + res.error);
    }
  }
</script>

<div class="execution-card {execution ? (execution.error ? 'card-error' : 'card-success') : 'card-running'}">
  <!-- 默认折叠状态栏 (简洁一两句话) -->
  <div class="summary-bar" on:click={() => (isExpanded = !isExpanded)}>
    <div class="status-left">
      {#if isExecuting}
        <Loader2 size={16} class="spinner" color="#0078D4" />
        <span class="summary-text font-running">正在生成并执行 Excel VBA...</span>
      {:else if execution?.error}
        <XCircle size={16} color="#A80000" />
        <span class="summary-text font-error" title={execution.summary}>
          {execution.summary || '执行遇到错误'}
        </span>
        {#if execution.elapsedMs}
          <span class="badge badge-gray">{(execution.elapsedMs / 1000).toFixed(2)}s</span>
        {/if}
      {:else}
        <CheckCircle2 size={16} color="#107C41" />
        <span class="summary-text font-success" title={execution?.summary}>
          {execution?.summary || '执行已完成'}
        </span>
        {#if execution?.elapsedMs}
          <span class="badge badge-green">{(execution.elapsedMs / 1000).toFixed(2)}s</span>
        {/if}
      {/if}
    </div>

    <button class="expand-btn" type="button">
      <span>{isExpanded ? '收起详情' : '展开代码与记录'}</span>
      {#if isExpanded}
        <ChevronUp size={14} />
      {:else}
        <ChevronDown size={14} />
      {/if}
    </button>
  </div>

  <!-- 展开后的详情面板 (默认完全折叠，点击才展开) -->
  {#if isExpanded && (execution || displayCode)}
    <div class="expanded-panel">
      <!-- 选项卡头部 -->
      <div class="panel-tabs">
        <div class="tabs-left">
          <button class="tab-btn {activeTab === 'code' ? 'active' : ''}" on:click={() => (activeTab = 'code')}>
            <Code size={13} />
            <span>VBA 源码</span>
          </button>
          <button class="tab-btn {activeTab === 'log' ? 'active' : ''}" on:click={() => (activeTab = 'log')}>
            <FileText size={13} />
            <span>执行日志</span>
          </button>
        </div>

        {#if activeTab === 'code'}
          <div class="tabs-actions">
            <button class="btn btn-sm" on:click={handleCopy} title="复制代码">
              {#if copySuccess}
                <Check size={12} color="#107C41" />
                <span style="color: #107C41;">已复制</span>
              {:else}
                <Copy size={12} />
                <span>复制</span>
              {/if}
            </button>
            <button class="btn btn-sm" on:click={() => (showSaveDialog = !showSaveDialog)} title="存入本地脚本目录">
              <Save size={12} />
              <span>保存到我的脚本</span>
            </button>
          </div>
        {/if}
      </div>

      <!-- Tab 1: 代码展示 -->
      {#if activeTab === 'code'}
        <div class="code-container">
          <pre class="vba-code"><code>{displayCode || '正在生成代码...'}</code></pre>
        </div>
      {/if}

      <!-- Tab 2: 日志展示 -->
      {#if activeTab === 'log'}
        <div class="log-container">
          {#if execution}
            <div class="log-row">
              <span class="log-label">耗时:</span>
              <span class="log-val">{execution.elapsedMs} ms</span>
            </div>
            <div class="log-row">
              <span class="log-label">指令摘要:</span>
              <span class="log-val">{prompt}</span>
            </div>
            {#if execution.error}
              <div class="log-row error-block">
                <span class="log-label">错误信息:</span>
                <pre class="error-text">{execution.error}</pre>
              </div>
            {:else}
              <div class="log-row">
                <span class="log-label">状态:</span>
                <span class="log-val" style="color: #107C41; font-weight: 500;">COM 调用顺利，模块已瞬时销毁清理</span>
              </div>
            {/if}
          {:else}
            <div class="log-row">
              <span class="log-label">状态:</span>
              <span class="log-val" style="color: #0078D4;">脚本执行准备中...</span>
            </div>
          {/if}
        </div>
      {/if}

      <!-- 保存到我的脚本内联弹窗 -->
      {#if showSaveDialog}
        <div class="save-dialog-inline">
          <div class="save-title">保存为本地独立脚本 (.bas)</div>
          <div class="save-row">
            <input type="text" placeholder="输入脚本名称 (如: 销售数据汇总)" bind:value={scriptName} />
            <button class="btn btn-primary btn-sm" on:click={handleSaveScript} disabled={isSaving}>
              {#if saveSuccess}
                <Check size={13} /> 已保存
              {:else if isSaving}
                保存中...
              {:else}
                确认保存
              {/if}
            </button>
          </div>
          <input
            type="text"
            class="desc-input"
            placeholder="可选：输入用途简述"
            bind:value={scriptDesc}
          />
        </div>
      {/if}

      <!-- 核心功能：多版本时间轴整本恢复控制栏 -->
      <div class="snapshot-footer">
        <div class="snap-info">
          <History size={14} color="#605E5C" />
          <span class="snap-label">恢复整本版本:</span>
          {#if allSnapshots.length > 0}
            <select class="snap-select" bind:value={selectedSnapshotId}>
              {#each allSnapshots as s}
                <option value={s.id}>
                  {s.timeDisplay} ({s.promptSummary.slice(0, 16)})
                </option>
              {/each}
            </select>
          {:else if execution.snapshot}
            <span class="snap-tag">{execution.snapshot.timeDisplay} 快照</span>
          {:else}
            <span class="snap-none">未生成快照</span>
          {/if}
        </div>

        {#if selectedSnapshotId || execution.snapshot}
          <button
            class="btn btn-sm btn-rollback"
            on:click={() => handleRestore(selectedSnapshotId || execution.snapshot?.id || '')}
            disabled={isRestoring}
          >
            {#if restoreSuccess}
              <Check size={13} color="#107C41" />
              <span style="color: #107C41; font-weight: 600;">已恢复整本工作簿！</span>
            {:else if isRestoring}
              <Loader2 size={13} class="spinner" /> 恢复中...
            {:else}
              <RotateCcw size={13} />
              <span>恢复到执行前</span>
            {/if}
          </button>
        {/if}
      </div>
    </div>
  {/if}
</div>

<style>
  .execution-card {
    background: #ffffff;
    border-radius: 6px;
    border: 1px solid var(--office-border);
    margin: 8px 0;
    overflow: hidden;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
    transition: all 0.15s ease;
  }

  .card-success {
    border-left: 3px solid var(--excel-green);
  }

  .card-error {
    border-left: 3px solid var(--office-danger);
  }

  .card-running {
    border-left: 3px solid var(--office-blue);
  }

  /* 默认一两句话简报栏 */
  .summary-bar {
    padding: 8px 12px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    cursor: pointer;
    background: #ffffff;
  }

  .summary-bar:hover {
    background: #faf9f8;
  }

  .status-left {
    display: flex;
    align-items: center;
    gap: 8px;
    min-width: 0;
    flex: 1;
  }

  .summary-text {
    font-size: 12px;
    font-weight: 500;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .font-success {
    color: #107c41;
  }

  .font-error {
    color: #a80000;
  }

  .font-running {
    color: #0078d4;
  }

  .expand-btn {
    display: flex;
    align-items: center;
    gap: 3px;
    font-size: 11px;
    color: var(--office-muted);
    background: none;
    border: none;
    cursor: pointer;
    padding: 2px 4px;
    border-radius: 3px;
    flex-shrink: 0;
  }

  .expand-btn:hover {
    background: #edebe9;
    color: var(--office-text);
  }

  /* 展开后面板 */
  .expanded-panel {
    border-top: 1px solid var(--office-border);
    background: #fbfbfb;
  }

  .panel-tabs {
    height: 32px;
    padding: 0 12px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 1px solid var(--office-border);
    background: #f3f2f1;
  }

  .tabs-left {
    display: flex;
    height: 100%;
  }

  .tab-btn {
    display: flex;
    align-items: center;
    gap: 4px;
    padding: 0 10px;
    font-size: 11px;
    font-weight: 500;
    color: var(--office-muted);
    background: transparent;
    border: none;
    border-bottom: 2px solid transparent;
    cursor: pointer;
    height: 100%;
  }

  .tab-btn.active {
    color: var(--excel-green);
    border-bottom-color: var(--excel-green);
    background: #ffffff;
    font-weight: 600;
  }

  .tabs-actions {
    display: flex;
    gap: 6px;
  }

  .btn-sm {
    padding: 2px 8px;
    font-size: 11px;
  }

  .code-container {
    padding: 10px 12px;
    background: #1e1e1e;
    color: #d4d4d4;
    max-height: 240px;
    overflow-y: auto;
  }

  .vba-code {
    font-family: Consolas, "Courier New", monospace;
    font-size: 11px;
    line-height: 1.5;
    white-space: pre-wrap;
    word-break: break-all;
  }

  .log-container {
    padding: 10px 12px;
    font-size: 11px;
    display: flex;
    flex-direction: column;
    gap: 6px;
    max-height: 200px;
    overflow-y: auto;
  }

  .log-row {
    display: flex;
    gap: 8px;
  }

  .log-label {
    color: var(--office-muted);
    min-width: 60px;
  }

  .log-val {
    color: var(--office-text);
    word-break: break-all;
  }

  .error-block {
    flex-direction: column;
    gap: 4px;
  }

  .error-text {
    background: var(--office-danger-light);
    color: var(--office-danger);
    padding: 6px 8px;
    border-radius: 4px;
    font-family: Consolas, monospace;
    font-size: 11px;
    white-space: pre-wrap;
  }

  /* 保存脚本内联面板 */
  .save-dialog-inline {
    padding: 8px 12px;
    background: #f3f9f5;
    border-bottom: 1px solid #d0ebd8;
    display: flex;
    flex-direction: column;
    gap: 6px;
  }

  .save-title {
    font-size: 11px;
    font-weight: 600;
    color: var(--excel-green);
  }

  .save-row {
    display: flex;
    gap: 6px;
  }

  .save-row input {
    flex: 1;
    height: 26px;
    padding: 0 8px;
    font-size: 11px;
    border: 1px solid #c2e2cc;
    border-radius: 3px;
    outline: none;
  }

  .desc-input {
    height: 24px;
    padding: 0 8px;
    font-size: 11px;
    border: 1px solid #e1e1e1;
    border-radius: 3px;
  }

  /* 快照回滚底栏 */
  .snapshot-footer {
    padding: 8px 12px;
    background: #ffffff;
    border-top: 1px solid var(--office-border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
  }

  .snap-info {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 11px;
    color: var(--office-muted);
    min-width: 0;
    flex: 1;
  }

  .snap-select {
    height: 24px;
    font-size: 11px;
    border: 1px solid var(--office-border);
    border-radius: 3px;
    padding: 0 4px;
    background: white;
    max-width: 170px;
  }

  .snap-tag {
    font-weight: 500;
    color: var(--office-text);
  }

  .btn-rollback {
    color: #795b00;
    background: #fff8e5;
    border-color: #f7e6b5;
    font-weight: 600;
    flex-shrink: 0;
  }

  .btn-rollback:hover {
    background: #ffefc4;
  }

  .spinner {
    animation: spin 1s linear infinite;
  }

  @keyframes spin {
    from {
      transform: rotate(0deg);
    }
    to {
      transform: rotate(360deg);
    }
  }
</style>
