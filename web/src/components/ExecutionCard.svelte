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
  let codeViewMode: 'executed' | 'original' | 'wrapper' = 'executed';

  $: executedCode = execution?.executedVbaCode || execution?.vbaCode || streamCode || '';
  $: originalCode = execution?.originalVbaCode || executedCode;
  $: wrapperCode = execution?.wrapperCode || '';
  $: displayCode =
    codeViewMode === 'original'
      ? originalCode
      : codeViewMode === 'wrapper'
      ? wrapperCode
      : executedCode;

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
    // 始终保存真实执行过、可再次运行的实际代码
    const codeToSave = executedCode || displayCode;
    if (!scriptName.trim() || !codeToSave) return;
    isSaving = true;
    const res = await bridge.send('save_script', {
      name: scriptName.trim(),
      code: codeToSave,
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
    if (
      !confirm(
        '【高危回滚提醒】\n确定将工作簿恢复到该宏执行前的快照吗？\n\n⚠️ 注意：宏执行后您所做的所有新增手工修改、单元格编辑都将被覆盖并永久丢失！\n（系统会在恢复前为当前状态保留一份紧急安全救援副本）'
      )
    ) {
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
        <div class="code-view-wrapper">
          <div class="code-sub-bar">
            <div class="code-sub-tabs">
              <button
                class="sub-tab-btn {codeViewMode === 'executed' ? 'active' : ''}"
                on:click={() => (codeViewMode = 'executed')}
              >
                实际执行源码
                {#if execution?.executedCodeHash}
                  <span class="hash-tag">SHA:{execution.executedCodeHash.slice(0, 8)}</span>
                {/if}
              </button>
              <button
                class="sub-tab-btn {codeViewMode === 'original' ? 'active' : ''}"
                on:click={() => (codeViewMode = 'original')}
              >
                模型原始提取
                {#if execution?.originalCodeHash}
                  <span class="hash-tag">SHA:{execution.originalCodeHash.slice(0, 8)}</span>
                {/if}
              </button>
              {#if wrapperCode}
                <button
                  class="sub-tab-btn {codeViewMode === 'wrapper' ? 'active' : ''}"
                  on:click={() => (codeViewMode = 'wrapper')}
                >
                  入口包装器
                </button>
              {/if}
            </div>

            <div class="code-integrity-badge">
              {#if execution?.isSourceIdentical}
                <span class="pill pill-green">正文与执行代码 100% 一致 (零暗改)</span>
              {:else if wrapperCode}
                <span class="pill pill-blue">正文零修改 + 追加透明包装器</span>
              {/if}
            </div>
          </div>

          <div class="code-container">
            <pre class="vba-code"><code>{displayCode || '正在生成代码...'}</code></pre>
          </div>
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

          <!-- 哈希与源码一致性审计 -->
          <div class="log-row">
            <span class="log-label">源码哈希:</span>
            <div class="hash-table">
              <div class="hash-item">
                <span class="hash-name">模型提取:</span>
                <code>{execution?.originalCodeHash || 'N/A'}</code>
              </div>
              <div class="hash-item">
                <span class="hash-name">实际执行:</span>
                <code>{execution?.executedCodeHash || 'N/A'}</code>
              </div>
              <div class="hash-status">
                {#if execution?.isSourceIdentical}
                  <span class="pill pill-green">哈希一致 (模型原貌直调)</span>
                {:else if wrapperCode}
                  <span class="pill pill-blue">正文零暗改 (追加受控入口包装器)</span>
                {/if}
              </div>
            </div>
          </div>

          {#if execution?.transformSteps && execution.transformSteps.length > 0}
            <div class="log-row flex-col">
              <span class="log-label">规整与包装:</span>
              <ul class="step-list">
                {#each execution.transformSteps as step}
                  <li>{step}</li>
                {/each}
              </ul>
            </div>
          {/if}

          <!-- 预编译与运行阶段状态精确区分 -->
          <div class="log-row">
            <span class="log-label">预检与阶段:</span>
            <div class="rb-badges">
              {#if execution?.precheckStatus === 'passed'}
                <span class="pill pill-green">静态预检通过 (VBE 578)</span>
              {:else if execution?.precheckStatus === 'unavailable'}
                <span class="pill pill-amber">预编译不可用 (未受检执行)</span>
              {:else if execution?.precheckStatus === 'failed'}
                <span class="pill pill-red">预编译拦截 (存在语法/引用错误)</span>
              {:else if execution?.precheckStatus === 'scope_risk_intercepted'}
                <span class="pill pill-red">范围失控拦截 (阻止运行)</span>
              {:else if execution?.precheckStatus === 'workbook_locked'}
                <span class="pill pill-red">工作簿已锁定保护</span>
              {/if}

              {#if execution?.executionPhase === 'macro_completed'}
                <span class="pill pill-blue">宏真正运行完成</span>
              {:else if execution?.executionPhase === 'hang_interrupted_recovered'}
                <span class="pill pill-amber">挂起已中断并证实恢复</span>
              {:else if execution?.executionPhase === 'hang_unconfirmed_locked'}
                <span class="pill pill-red">挂起中断未证实 (已锁定)</span>
              {:else if execution?.executionPhase === 'hang_suspected_interrupt_sent'}
                <span class="pill pill-amber">疑似挂起已发中断</span>
              {:else if execution?.executionPhase === 'runtime_hang'}
                <span class="pill pill-red">运行期挂起</span>
              {:else if execution?.executionPhase === 'runtime_error'}
                <span class="pill pill-red">运行期异常抛出</span>
              {:else if execution?.executionPhase === 'syntax_failed' || execution?.executionPhase === 'intercepted_before_run' || execution?.executionPhase === 'blocked_by_lock'}
                <span class="pill pill-gray">未进入运行阶段</span>
              {/if}
            </div>
          </div>

          <!-- 生成重试与 Token 成本明细 (纠正标签数学逻辑) -->
          {#if execution?.retryCount !== undefined || execution?.llmCost}
            <div class="log-row">
              <span class="log-label">生成开销:</span>
              <div class="cost-info">
                <span class="cost-item">
                  重试: <strong>{execution?.retryCount || 0} 次</strong>
                  {#if (execution?.retryCount || 0) > 0}
                    <span class="dim-text">(安全纠偏/重试，绝不拼装半截代码)</span>
                  {/if}
                </span>
                {#if execution?.llmCost?.totalTokens}
                  {@const reasoningTokens = execution.llmCost.reasoningTokens || 0}
                  {@const compTokens = execution.llmCost.completionTokens || 0}
                  {@const contentTokens = execution.llmCost.contentTokens ?? (compTokens >= reasoningTokens ? compTokens - reasoningTokens : compTokens)}
                  <span class="cost-item">
                    Token 累计: <strong>{execution.llmCost.totalTokens}</strong>
                    {#if reasoningTokens > 0}
                      <span class="dim-text">(生成={compTokens}: 思考={reasoningTokens} + 正文={contentTokens})</span>
                    {:else if contentTokens > 0}
                      <span class="dim-text">(正文: {contentTokens})</span>
                    {/if}
                  </span>
                  {#if execution.llmCost.promptTokens}
                    <span class="cost-item dim-text">提示词: {execution.llmCost.promptTokens}</span>
                  {/if}
                {/if}
              </div>
            </div>
          {/if}

          {#if execution?.riskNotice}
            <div class="log-row flex-col">
              <span class="log-label">运行状态说明:</span>
              <span class="log-val dim-text">{execution.riskNotice}</span>
            </div>
          {/if}

          <!-- 跨工作簿隔离检查 -->
          <div class="log-row">
            <span class="log-label">跨文件检查:</span>
            {#if execution?.readback?.otherWorkbooksAffected}
              <span class="log-val font-error">
                ⚠️ 检测到其他打开的工作簿受到影响：{execution.readback.affectedWorkbooksWarning}
              </span>
            {:else}
              <span class="log-val" style="color: #107C41;">未影响其他已打开工作簿</span>
            {/if}
          </div>

          {#if execution?.error}
            <div class="log-row error-block">
              <span class="log-label">异常详情:</span>
              <pre class="error-text">{execution.error}</pre>
            </div>
          {:else}
            <div class="log-row">
              <span class="log-label">模块清理:</span>
              <span class="log-val" style="color: #107C41;">COM 宏调用完成，临时模块已瞬时销毁清理</span>
            </div>
          {/if}

          <!-- 开放式 VBA 安全与回滚边界警示 -->
          <div class="boundary-warning-card">
            <div class="bw-header">
              <ShieldAlert size={14} color="#D83B01" />
              <span class="bw-title">开放式 VBA 安全与回滚边界声明</span>
            </div>
            <p class="bw-body">
              工作簿快照仅保障目标工作簿本身的数据与格式原位回滚。若模型生成的开放式 VBA 包含外部文件读写、修改了其他工作簿、调用系统 API 或执行外部进程等副作用，快照无法自动撤回。宿主负责核验目标并确保单点执行，宏的系统级行为仍需人工把关。
            </p>
          </div>
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

      <!-- 快照保护边界明确声明，不夸大快照覆盖范围 -->
      <div class="snapshot-boundary-note">
        <span>ℹ️ 快照保护边界：快照完整备份和回滚目标工作簿本身的数据与结构；VBA 宏若涉及外部工作簿、本地磁盘或网络操作等外部副作用，快照无法自动回滚。</span>
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

  .pill-amber {
    background: #fff4ce;
    color: #797673;
  }

  .cost-info {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 8px;
    font-size: 11px;
    color: var(--office-text);
  }

  .cost-item {
    display: inline-flex;
    align-items: center;
    gap: 4px;
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

  .code-view-wrapper {
    display: flex;
    flex-direction: column;
  }

  .code-sub-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #252526;
    padding: 4px 10px;
    border-bottom: 1px solid #333333;
  }

  .code-sub-tabs {
    display: flex;
    gap: 4px;
  }

  .sub-tab-btn {
    background: transparent;
    border: 1px solid transparent;
    color: #969696;
    font-size: 10px;
    padding: 2px 7px;
    border-radius: 3px;
    cursor: pointer;
    display: flex;
    align-items: center;
    gap: 4px;
  }

  .sub-tab-btn:hover {
    color: #e0e0e0;
    background: #2d2d2d;
  }

  .sub-tab-btn.active {
    color: #ffffff;
    background: #37373d;
    border-color: #4b4b4b;
    font-weight: 600;
  }

  .hash-tag {
    font-family: Consolas, monospace;
    font-size: 9px;
    color: #4ec9b0;
    opacity: 0.85;
  }

  .code-integrity-badge {
    display: flex;
    align-items: center;
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

  .hash-table {
    display: flex;
    flex-direction: column;
    gap: 4px;
    background: #f3f2f1;
    padding: 6px 8px;
    border-radius: 4px;
    flex: 1;
  }

  .hash-item {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 10px;
  }

  .hash-name {
    color: var(--office-muted);
    min-width: 60px;
  }

  .hash-item code {
    font-family: Consolas, monospace;
    font-size: 10px;
    color: #004e8c;
    word-break: break-all;
  }

  .hash-status {
    margin-top: 2px;
  }

  .boundary-warning-card {
    margin-top: 6px;
    padding: 8px 10px;
    background: #fff8f5;
    border: 1px solid #fed9cc;
    border-radius: 4px;
  }

  .bw-header {
    display: flex;
    align-items: center;
    gap: 5px;
    margin-bottom: 4px;
  }

  .bw-title {
    font-size: 11px;
    font-weight: 600;
    color: #d83b01;
  }

  .bw-body {
    margin: 0;
    font-size: 10.5px;
    line-height: 1.45;
    color: #605e5c;
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

  .snapshot-boundary-note {
    font-size: 11px;
    color: var(--office-muted);
    padding: 6px 12px;
    background: #fbfbfb;
    border-top: 1px dashed var(--office-border);
    line-height: 1.4;
  }
</style>
