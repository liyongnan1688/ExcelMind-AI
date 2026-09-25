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
    SearchCheck,
    ShieldAlert,
    AlertCircle,
  } from 'lucide-svelte';
  import { bridge, type SnapshotItem, type VbaExecutionData } from '../services/bridge';

  export let prompt: string;
  export let execution: VbaExecutionData | null = null;
  export let streamCode: string = '';
  export let isExecuting = false;
  export let allSnapshots: SnapshotItem[] = [];
  export let onSaveScriptSuccess: () => void;

  // 严格默认折叠，绝不主动展开，保证对话界面清爽
  let isExpanded = false;
  let activeTab: 'code' | 'readback' | 'audit' = 'readback';

  $: displayCode = execution?.executedVbaCode || execution?.vbaCode || streamCode || '';
  $: rawCode = execution?.originalVbaCode || displayCode;

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
    const res = await bridge.send('restore_snapshot', {
      snapshotId: targetId,
      targetWorkbookName: execution?.targetWorkbookName || '',
    });
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

<div
  class="execution-card {isExecuting
    ? 'card-running'
    : execution?.error
    ? 'card-error'
    : execution?.verificationStatus === 'verified'
    ? 'card-success'
    : 'card-unconfirmed'}"
>
  <!-- 默认折叠状态栏 (简洁一两句话，客观如实呈现，不伪称满分完成) -->
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
      {:else if execution?.verificationStatus === 'verified'}
        <CheckCircle2 size={16} color="#107C41" />
        <span class="summary-text font-success" title={execution.summary}>
          {execution.summary || '宏已运行，区域验证通过'}
        </span>
        {#if execution?.elapsedMs}
          <span class="badge badge-green">{(execution.elapsedMs / 1000).toFixed(2)}s</span>
        {/if}
      {:else}
        <!-- 待确认或有差异状态 (显示中性/提示色，不伪装绿色的“任务完成”) -->
        <AlertCircle size={16} color="#D83B01" />
        <span class="summary-text font-unconfirmed" title={execution?.summary}>
          {execution?.summary || '宏已运行，效果待确认'}
        </span>
        <span class="badge badge-amber">效果待确认</span>
        {#if execution?.elapsedMs}
          <span class="badge badge-gray">{(execution.elapsedMs / 1000).toFixed(2)}s</span>
        {/if}
      {/if}
    </div>

    <button class="expand-btn" type="button">
      <span>{isExpanded ? '收起详情' : '展开核验与记录'}</span>
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
          <button
            class="tab-btn {activeTab === 'readback' ? 'active' : ''}"
            on:click={() => (activeTab = 'readback')}
          >
            <SearchCheck size={13} />
            <span>写后核验</span>
          </button>
          <button class="tab-btn {activeTab === 'code' ? 'active' : ''}" on:click={() => (activeTab = 'code')}>
            <Code size={13} />
            <span>VBA 源码</span>
          </button>
          <button class="tab-btn {activeTab === 'audit' ? 'active' : ''}" on:click={() => (activeTab = 'audit')}>
            <FileText size={13} />
            <span>审计日志</span>
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

      <!-- Tab 1: 写后核验详情 (从 Excel 真实读回的数据) -->
      {#if activeTab === 'readback'}
        <div class="readback-container">
          {#if execution?.readback}
            <div class="rb-row">
              <span class="rb-label">目标工作簿:</span>
              <span class="rb-val">
                {execution.readback.targetWorkbookName}
                {#if execution.readback.targetVerified}
                  <span class="pill pill-green">身份核验一致</span>
                {:else}
                  <span class="pill pill-red">身份不一致</span>
                {/if}
              </span>
            </div>

            <div class="rb-row">
              <span class="rb-label">目标工作表:</span>
              <span class="rb-val">{execution.readback.targetSheetName || '默认活动表'}</span>
            </div>

            <div class="rb-row">
              <span class="rb-label">实际使用区域:</span>
              <span class="rb-val highlight-val">
                {execution.readback.usedRangeAddress || '未检测到使用区域'}
                {#if execution.readback.rowCount > 0}
                  <span class="dim-text">
                    ({execution.readback.rowCount}行 × {execution.readback.columnCount}列，起始: {execution.readback.startCell || 'N/A'})
                  </span>
                {/if}
              </span>
            </div>

            <div class="rb-row">
              <span class="rb-label">样式与特征:</span>
              <div class="rb-badges">
                {#if execution.readback.hasBorders}
                  <span class="pill pill-blue">包含边框</span>
                {/if}
                {#if execution.readback.hasInteriorColor}
                  <span class="pill pill-blue">包含单元格背景色</span>
                {/if}
                {#if execution.readback.hasFormulas}
                  <span class="pill pill-blue">包含公式计算</span>
                {/if}
                {#if !execution.readback.hasBorders && !execution.readback.hasInteriorColor && !execution.readback.hasFormulas}
                  <span class="pill pill-gray">无特殊样式/纯文本填入</span>
                {/if}
              </div>
            </div>

            {#if execution.readback.sampleValues && execution.readback.sampleValues.length > 0}
              <div class="rb-row rb-samples">
                <span class="rb-label">单元格抽样:</span>
                <div class="sample-tags">
                  {#each execution.readback.sampleValues.slice(0, 8) as sample}
                    <code class="sample-code">{sample}</code>
                  {/each}
                </div>
              </div>
            {/if}

            {#if execution.verificationNote}
              <div class="verification-box">
                <span class="ver-label">核验摘要:</span>
                <span class="ver-text">{execution.verificationNote}</span>
              </div>
            {/if}
          {:else}
            <div class="dim-empty">尚未获取到写后读回数据</div>
          {/if}
        </div>
      {/if}

      <!-- Tab 2: 代码展示 -->
      {#if activeTab === 'code'}
        <div class="code-container">
          <pre class="vba-code"><code>{displayCode || '正在生成代码...'}</code></pre>
        </div>
      {/if}

      <!-- Tab 3: 审计日志 -->
      {#if activeTab === 'audit'}
        <div class="log-container">
          <div class="log-row">
            <span class="log-label">用户指令:</span>
            <span class="log-val">{prompt}</span>
          </div>
          <div class="log-row">
            <span class="log-label">绑定目标:</span>
            <span class="log-val">{execution?.targetWorkbookName || '当前活动工作簿'}</span>
          </div>
          <div class="log-row">
            <span class="log-label">调用耗时:</span>
            <span class="log-val">{execution?.elapsedMs || 0} ms</span>
          </div>

          {#if execution?.transformSteps && execution.transformSteps.length > 0}
            <div class="log-row flex-col">
              <span class="log-label">宿主规整步骤:</span>
              <ul class="step-list">
                {#each execution.transformSteps as step}
                  <li>{step}</li>
                {/each}
              </ul>
            </div>
          {/if}

          {#if execution?.error}
            <div class="log-row error-block">
              <span class="log-label">异常详情:</span>
              <pre class="error-text">{execution.error}</pre>
            </div>
          {:else}
            <div class="log-row">
              <span class="log-label">运行状态:</span>
              <span class="log-val" style="color: #107C41;">COM 宏调用完成，临时模块已瞬时销毁清理</span>
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
          <span class="snap-label">整本版本回滚:</span>
          {#if allSnapshots.length > 0}
            <select class="snap-select" bind:value={selectedSnapshotId}>
              {#each allSnapshots as s}
                <option value={s.id}>
                  {s.timeDisplay} ({s.promptSummary.slice(0, 16)})
                </option>
              {/each}
            </select>
          {:else if execution?.snapshot?.timeDisplay}
            <span class="snap-tag">{execution.snapshot.timeDisplay} 快照</span>
          {:else}
            <span class="snap-none">未生成快照 (未保存工作簿)</span>
          {/if}
        </div>

        {#if selectedSnapshotId || execution?.snapshot?.id}
          <button
            class="btn btn-sm btn-rollback"
            on:click={() => handleRestore(selectedSnapshotId || execution?.snapshot?.id || '')}
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

  .card-unconfirmed {
    border-left: 3px solid #d83b01;
  }

  .card-error {
    border-left: 3px solid var(--office-danger);
  }

  .card-running {
    border-left: 3px solid var(--office-blue);
  }

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

  .font-unconfirmed {
    color: #d83b01;
  }

  .font-error {
    color: #a80000;
  }

  .font-running {
    color: #0078d4;
  }

  .badge-amber {
    background: #fdf3eb;
    color: #d83b01;
    border: 1px solid #fed9cc;
    font-size: 10px;
    padding: 1px 5px;
    border-radius: 3px;
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

  /* 写后核验面板样式 */
  .readback-container {
    padding: 10px 12px;
    font-size: 11px;
    display: flex;
    flex-direction: column;
    gap: 6px;
    background: #ffffff;
  }

  .rb-row {
    display: flex;
    align-items: baseline;
    gap: 8px;
  }

  .rb-label {
    color: var(--office-muted);
    min-width: 80px;
    flex-shrink: 0;
  }

  .rb-val {
    color: var(--office-text);
    word-break: break-all;
    display: flex;
    align-items: center;
    gap: 6px;
  }

  .highlight-val {
    font-weight: 600;
    color: #107c41;
  }

  .dim-text {
    font-weight: normal;
    color: var(--office-muted);
    font-size: 11px;
  }

  .pill {
    padding: 1px 6px;
    border-radius: 3px;
    font-size: 10px;
    font-weight: 500;
  }

  .pill-green {
    background: #e7f3ec;
    color: #107c41;
  }

  .pill-blue {
    background: #eff6fc;
    color: #0078d4;
  }

  .pill-red {
    background: #fdf3f4;
    color: #a80000;
  }

  .pill-gray {
    background: #f3f2f1;
    color: #605e5c;
  }

  .sample-tags {
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
  }

  .sample-code {
    background: #f3f2f1;
    padding: 2px 5px;
    border-radius: 3px;
    font-family: Consolas, monospace;
    font-size: 10px;
    color: #201f1e;
  }

  .verification-box {
    margin-top: 4px;
    padding: 6px 8px;
    background: #fcf9f5;
    border: 1px solid #fae8d4;
    border-radius: 4px;
    display: flex;
    gap: 6px;
    font-size: 11px;
  }

  .ver-label {
    color: #d83b01;
    font-weight: 600;
    flex-shrink: 0;
  }

  .ver-text {
    color: #323130;
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
    max-height: 220px;
    overflow-y: auto;
  }

  .log-row {
    display: flex;
    gap: 8px;
  }

  .flex-col {
    flex-direction: column;
    gap: 3px;
  }

  .log-label {
    color: var(--office-muted);
    min-width: 80px;
  }

  .log-val {
    color: var(--office-text);
    word-break: break-all;
  }

  .step-list {
    margin: 2px 0 0 16px;
    padding: 0;
    color: #323130;
  }

  .step-list li {
    margin-bottom: 2px;
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

  .dim-empty {
    color: var(--office-muted);
    font-style: italic;
    padding: 4px 0;
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
