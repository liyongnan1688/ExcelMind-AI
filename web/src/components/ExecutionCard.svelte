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
    Download,
    Maximize2,
    Minimize2,
    X,
  } from 'lucide-svelte';
  import { bridge, type SnapshotItem, type VbaExecutionData } from '../services/bridge';
  import { getPromptConstraintsCatalog } from '../services/llm';
  import { createEventDispatcher, tick } from 'svelte';

  const dispatch = createEventDispatcher();
  const promptConstraints = getPromptConstraintsCatalog();
  let showConstraints = false;

  export let prompt: string;
  export let execution: VbaExecutionData | null = null;
  export let streamCode: string = '';
  export let isExecuting = false;
  export let allSnapshots: SnapshotItem[] = [];
  export let onSaveScriptSuccess: () => void;

  let cardElement: HTMLElement;

  // 严格默认折叠，绝不主动展开，保证对话界面清爽
  let isExpanded = false;
  let isFullscreen = false;
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

  // 折叠/展开切换：展开时自动平滑滚动对齐，彻底解决“点开显示不全”
  async function toggleExpand() {
    isExpanded = !isExpanded;
    if (isExpanded) {
      await tick();
      dispatch('expand');
      if (cardElement) {
        cardElement.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }
    }
  }

  // 双重稳健复制逻辑 (支持 Clipboard API + execCommand 降级兜底)
  let copyActiveType: 'current' | 'original' | 'executed' | 'wrapper' | null = null;
  let copySuccess = false;
  let copyFeedbackText = '';

  async function copyTextToClipboard(text: string, type: 'current' | 'original' | 'executed' | 'wrapper', label: string) {
    if (!text) return;
    let ok = false;
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(text);
        ok = true;
      }
    } catch (e) {}

    if (!ok) {
      try {
        const ta = document.createElement('textarea');
        ta.value = text;
        ta.style.position = 'fixed';
        ta.style.opacity = '0';
        ta.style.left = '-9999px';
        document.body.appendChild(ta);
        ta.focus();
        ta.select();
        ok = document.execCommand('copy');
        document.body.removeChild(ta);
      } catch (e) {}
    }

    if (ok) {
      copyActiveType = type;
      copySuccess = true;
      copyFeedbackText = `已复制${label}`;
      setTimeout(() => {
        copySuccess = false;
        copyActiveType = null;
        copyFeedbackText = '';
      }, 1800);
    }
  }

  function handleCopyCurrent() {
    const label = codeViewMode === 'original' ? '原始提取源码' : (codeViewMode === 'wrapper' ? '入口包装器' : '校验执行源码');
    copyTextToClipboard(displayCode, 'current', label);
  }

  function handleCopyOriginal() {
    copyTextToClipboard(originalCode, 'original', '原始提取源码');
  }

  function handleCopyExecuted() {
    copyTextToClipboard(executedCode, 'executed', '校验执行源码');
  }

  function handleCopyWrapper() {
    copyTextToClipboard(wrapperCode, 'wrapper', '入口包装器');
  }

  function handleExportVba(targetType: 'original' | 'executed' | 'wrapper') {
    const text = targetType === 'original' ? originalCode : (targetType === 'wrapper' ? wrapperCode : executedCode);
    if (!text) return;
    const namePrefix = targetType === 'original' ? 'model_original_extract' : (targetType === 'wrapper' ? 'entry_wrapper' : 'executed_vba_runner');
    const filename = `${namePrefix}_${Date.now()}.vba`;
    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
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

  let showRestoreConfirm = false;
  let restoreTargetId = '';
  let restoreErrorMsg = '';

  function openRestoreConfirm(targetId: string) {
    if (!targetId) return;
    restoreTargetId = targetId;
    restoreErrorMsg = '';
    showRestoreConfirm = true;
  }

  async function confirmRestore() {
    if (!restoreTargetId) return;
    isRestoring = true;
    showRestoreConfirm = false;
    const res = await bridge.send('restore_snapshot', {
      snapshotId: restoreTargetId,
      targetWorkbookName: execution?.targetWorkbookName || '',
      targetWorkbookFullName: execution?.targetWorkbookFullName || execution?.readback?.targetWorkbookFullName || '',
    });
    isRestoring = false;

    if (res.ok) {
      restoreSuccess = true;
      dispatch('restored', {
        snapshotId: restoreTargetId,
        targetWorkbookName: execution?.targetWorkbookName,
        targetWorkbookFullName: execution?.targetWorkbookFullName || execution?.readback?.targetWorkbookFullName,
      });
      setTimeout(() => {
        restoreSuccess = false;
      }, 2500);
    } else {
      restoreErrorMsg = res.error || '恢复操作遇到错误';
    }
  }
</script>

<div
  bind:this={cardElement}
  class="execution-card {isExecuting
    ? 'card-running'
    : execution?.error
    ? 'card-error'
    : execution?.verificationStatus === 'verified'
    ? 'card-success'
    : 'card-unconfirmed'}"
>
  <!-- 默认折叠状态栏 (简洁呈现，客观如实呈现，不伪称满分完成) -->
  <div
    class="summary-bar"
    role="button"
    tabindex="0"
    on:click={toggleExpand}
    on:keydown={(e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        toggleExpand();
        e.preventDefault();
      }
    }}
  >
    <div class="status-left">
      {#if isExecuting}
        <Loader2 size={16} class="spinner" color="#0078D4" />
        <span class="summary-text font-running">正在生成并执行 Excel VBA...</span>
      {:else if execution?.error}
        <XCircle size={16} color="#C42B1C" />
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
        <AlertCircle size={16} color="#B25900" />
        <span class="summary-text font-unconfirmed" title={execution?.summary}>
          {execution?.summary || '宏已运行，效果待确认'}
        </span>
        <span class="badge badge-amber">效果待确认</span>
        {#if execution?.elapsedMs}
          <span class="badge badge-gray">{(execution.elapsedMs / 1000).toFixed(2)}s</span>
        {/if}
      {/if}
    </div>

    <div class="summary-right-actions">
      <!-- 双模态查看：未展开时也可直接全屏放大查看源码与核验 -->
      {#if execution || displayCode}
        <button
          class="btn-icon-subtle"
          type="button"
          title="全屏放大查看源码与核验报告"
          on:click|stopPropagation={() => (isFullscreen = true)}
        >
          <Maximize2 size={13} />
        </button>
      {/if}
      <button class="expand-btn" type="button" aria-expanded={isExpanded}>
        <span>{isExpanded ? '收起详情' : '展开核验与记录'}</span>
        {#if isExpanded}
          <ChevronUp size={14} />
        {:else}
          <ChevronDown size={14} />
        {/if}
      </button>
    </div>
  </div>

  <!-- 部分修改常驻告警栏 (显式告知部分成功及已影响区域，并提供一键整本恢复) -->
  {#if execution?.isPartiallyModified}
    <div class="partial-mod-banner">
      <div class="partial-mod-left">
        <AlertCircle size={16} color="#B25900" />
        <div class="partial-mod-text">
          <span class="partial-mod-title">宏运行失败，工作簿可能已部分修改</span>
          <span class="partial-mod-desc">
            执行后读回区域: <strong>{execution?.readback?.usedRangeAddress || '见读回'}</strong>；宏可能已部分修改工作簿。执行前快照已就绪，可随时整本恢复。
          </span>
        </div>
      </div>
      <div class="partial-mod-right">
        <button
          class="btn btn-sm btn-rollback"
          on:click|stopPropagation={() => openRestoreConfirm(selectedSnapshotId || execution?.snapshot?.id || '')}
        >
          <RotateCcw size={12} />
          <span>整本恢复快照</span>
        </button>
      </div>
    </div>
  {/if}

  <!-- 展开后的详情面板 (默认完全折叠，点击才展开) -->
  {#if isExpanded && (execution || displayCode)}
    <div class="expanded-panel">
      <!-- 选项卡头部与自适应工具栏 -->
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

        <div class="tabs-actions">
          {#if activeTab === 'code'}
            <button class="btn btn-sm" on:click={handleCopyCurrent} title="一键复制当前展示代码">
              {#if copyActiveType === 'current' || (copySuccess && !copyActiveType)}
                <Check size={12} color="#107C41" />
                <span class="copy-success-text">{copyFeedbackText || '已复制'}</span>
              {:else}
                <Copy size={12} />
                <span class="btn-label-text">复制当前</span>
              {/if}
            </button>
            <button class="btn btn-sm" on:click={() => (showSaveDialog = !showSaveDialog)} title="存入本地脚本目录">
              <Save size={12} />
              <span class="btn-label-text">存入脚本库</span>
            </button>
          {/if}
          <!-- 双模态查看全屏放大按钮 -->
          <button class="btn btn-sm btn-maximize" on:click={() => (isFullscreen = true)} title="全屏放大查看完整详情">
            <Maximize2 size={12} />
            <span class="btn-label-text">放大查看</span>
          </button>
        </div>
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
                title="查看校验并实际执行的 VBA 源码"
              >
                实际执行源码
                {#if execution?.executedCodeHash}
                  <span class="hash-tag">SHA:{execution.executedCodeHash.slice(0, 8)}</span>
                {/if}
              </button>
              <button
                class="sub-tab-btn {codeViewMode === 'original' ? 'active' : ''}"
                on:click={() => (codeViewMode = 'original')}
                title="查看大模型原始提取的代码（未经任何包装改动）"
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
                  title="查看自动注入的受控入口包装器"
                >
                  入口包装器
                </button>
              {/if}
            </div>

            <!-- 分别复制提取与完整性状态栏 -->
            <div class="code-sub-actions">
              <!-- 一键复制原始提取源码 -->
              <button
                class="sub-action-btn {copyActiveType === 'original' ? 'copied' : ''}"
                on:click={handleCopyOriginal}
                title="一键复制模型原始提取的代码"
              >
                {#if copyActiveType === 'original'}
                  <Check size={11} color="#4ec9b0" />
                  <span class="action-text-copied">已复制原始</span>
                {:else}
                  <Copy size={11} />
                  <span>复制原始</span>
                {/if}
              </button>

              <!-- 一键复制校验后执行源码 -->
              <button
                class="sub-action-btn {copyActiveType === 'executed' ? 'copied' : ''}"
                on:click={handleCopyExecuted}
                title="一键复制校验后实际执行的完整源码"
              >
                {#if copyActiveType === 'executed'}
                  <Check size={11} color="#4ec9b0" />
                  <span class="action-text-copied">已复制执行</span>
                {:else}
                  <Copy size={11} />
                  <span>复制执行</span>
                {/if}
              </button>

              <!-- 一键导出/提取 .vba 文件 -->
              <button
                class="sub-action-btn export-btn"
                on:click={() => handleExportVba(codeViewMode)}
                title="提取并下载当前代码为 .vba 文件"
              >
                <Download size={11} />
                <span>提取.vba</span>
              </button>

              <div class="code-integrity-badge">
                {#if execution?.isSourceIdentical}
                  <span class="pill pill-green">代码 100% 直跑一致</span>
                {:else if wrapperCode}
                  <span class="pill pill-blue">正文零修改 + 透明包装器</span>
                {/if}
              </div>
            </div>
          </div>

          <div class="code-container">
            <button
              class="floating-copy-btn {copyActiveType === 'current' ? 'copied' : ''}"
              on:click={handleCopyCurrent}
              title="一键复制当前窗口展示的代码"
            >
              {#if copyActiveType === 'current'}
                <Check size={12} color="#4ec9b0" />
                <span class="action-text-copied">已复制代码</span>
              {:else}
                <Copy size={12} />
                <span>复制代码</span>
              {/if}
            </button>
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

          <!-- API 请求与预算审计 -->
          <div class="log-row flex-col audit-box">
            <div class="audit-box-title">API 传输与输出预算审计:</div>
            <div class="audit-box-content">
              <div>• <strong>输出预算 (max_tokens):</strong> <code>{execution?.apiAudit?.maxTokensStatus || '未发送（采用服务端默认限制）'}</code></div>
              <div>• <strong>思考预算 (budget_tokens):</strong> <code>{execution?.apiAudit?.thinkingBudgetStatus || '未发送'}</code></div>
              <div>• <strong>发送历史条数:</strong> <code>{execution?.apiAudit?.historyCountSent ?? 'N/A'} 条</code>（可用总数: {execution?.apiAudit?.totalHistoryAvailable ?? 'N/A'} 条）</div>
              <div>• <strong>历史剥离/截断:</strong> <span class="pill pill-green">否（100% 完整保留含 VBA 的对话原文）</span></div>
              <div>• <strong>上下文策略:</strong> {execution?.apiAudit?.historyStrategy || '保留完整原文供模型接续调试'}</div>
            </div>
          </div>

          <!-- 系统提示词规则分类清单 -->
          <div class="log-row flex-col audit-box">
            <div class="audit-box-header">
              <span class="audit-box-title">系统提示词约束分类清单:</span>
              <button
                class="btn btn-sm"
                on:click={() => (showConstraints = !showConstraints)}
              >
                {showConstraints ? '收起清单' : '查看完整规则 (4类)'}
              </button>
            </div>
            {#if showConstraints}
              <div class="audit-constraints-list">
                {#each promptConstraints as rule}
                  <div class="constraint-item">
                    <span class="pill {rule.category === '模型能力限制' ? 'pill-red' : (rule.category === '运行入口协议' ? 'pill-blue' : 'pill-gray')}">
                      {rule.category}
                    </span>
                    <strong> {rule.name}:</strong> {rule.description}
                    {#if rule.isHostCapabilityLimitation}
                      <div class="constraint-warning">
                        ⚠️ {rule.limitationReason}
                      </div>
                    {/if}
                  </div>
                {/each}
              </div>
            {/if}
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
              {:else if execution?.precheckStatus === 'warning'}
                <span class="pill pill-amber">预检未确认 (整本快照保护下运行)</span>
              {:else if execution?.precheckStatus === 'unavailable'}
                <span class="pill pill-amber">预编译不可用 (整本快照保护下运行)</span>
              {:else if execution?.precheckStatus === 'failed'}
                <span class="pill pill-red">预编译拦截 (编译错误)</span>
              {:else if execution?.precheckStatus === 'scope_risk_intercepted'}
                <span class="pill pill-red">范围失控拦截 (阻止运行)</span>
              {:else if execution?.precheckStatus === 'workbook_locked'}
                <span class="pill pill-red">工作簿已锁定保护</span>
              {/if}

              {#if execution?.executionPhase === 'macro_completed'}
                <span class="pill pill-blue">宏真正运行完成</span>
              {:else if execution?.executionPhase === 'extract_failed'}
                <span class="pill pill-red">代码提取失败 (无有效宏)</span>
              {:else if execution?.executionPhase === 'injection_failed'}
                <span class="pill pill-red">模块注入失败 (权限/保护)</span>
              {:else if execution?.executionPhase === 'compile_failed' || execution?.executionPhase === 'syntax_failed'}
                <span class="pill pill-amber">静态编译未通过</span>
              {:else if execution?.executionPhase === 'invocation_failed'}
                <span class="pill pill-red">入口调度失败</span>
              {:else if execution?.executionPhase === 'runtime_error'}
                {#if execution?.isPartiallyModified}
                  <span class="pill pill-amber">运行异常 (工作簿已部分修改)</span>
                {:else}
                  <span class="pill pill-red">运行期异常抛出 ({execution?.vbaErrNumber ? 'Err ' + execution.vbaErrNumber : (execution?.rawErrorCode || '1004')})</span>
                {/if}
              {:else if execution?.executionPhase === 'hang_interrupted_recovered'}
                <span class="pill pill-amber">挂起已中断并证实恢复</span>
              {:else if execution?.executionPhase === 'hang_unconfirmed_locked'}
                <span class="pill pill-red">挂起中断未证实 (已锁定)</span>
              {:else if execution?.executionPhase === 'hang_suspected_interrupt_sent'}
                <span class="pill pill-amber">疑似挂起已发中断</span>
              {:else if execution?.executionPhase === 'runtime_hang'}
                <span class="pill pill-red">运行期挂起</span>
              {:else if execution?.executionPhase === 'intercepted_before_run' || execution?.executionPhase === 'blocked_by_lock'}
                <span class="pill pill-gray">未进入运行阶段</span>
              {/if}
            </div>
          </div>

          <!-- 完整错误链详细审计 -->
          {#if execution?.vbaErrNumber || execution?.vbaErrDescription || execution?.comHResult}
            <div class="log-row flex-col error-chain-box">
              <span class="log-label error-chain-label">完整错误链审计:</span>
              <div class="error-chain-content">
                {#if execution.vbaErrNumber}
                  <div>• <strong>VBA 运行时错误号:</strong> <code class="error-code">Err.Number = {execution.vbaErrNumber}</code></div>
                {/if}
                {#if execution.vbaErrDescription}
                  <div>• <strong>VBE 弹窗错误文本:</strong> <code>{execution.vbaErrDescription}</code></div>
                {/if}
                {#if execution.comHResult}
                  <div>• <strong>COM 宿主接口 HResult:</strong> <code>{execution.comHResult}</code> (由 app.Run 终止返回)</div>
                {/if}
                {#if execution.hostExecutionPhase}
                  <div>• <strong>宿主记录阶段:</strong> <span class="pill pill-red">{execution.hostExecutionPhase}</span></div>
                {/if}
              </div>
            </div>
          {/if}

          <!-- 工作簿部分修改与快照回滚提示 -->
          {#if execution?.isPartiallyModified}
            <div class="log-row flex-col partial-warning-box">
              <div class="partial-warning-header">
                <span class="partial-warning-title">⚠️ 宏运行异常中断：工作簿已发生部分修改！</span>
                {#if selectedSnapshotId || execution?.snapshot?.id}
                  <button
                    class="btn btn-sm btn-danger"
                    on:click={() => openRestoreConfirm(selectedSnapshotId || execution?.snapshot?.id || '')}
                  >
                    一键整本回滚
                  </button>
                {/if}
              </div>
              <div class="partial-warning-desc">
                执行后读回区域: <strong>{execution?.readback?.usedRangeAddress || '见读回'}</strong>；宏可能已部分修改工作簿。可点击上方按钮一键恢复至初始快照。
              </div>
            </div>
          {/if}

          {#if execution?.errorTriggerPoint}
            <div class="log-row">
              <span class="log-label">触发排查:</span>
              <span class="log-val trigger-point-val">{execution.errorTriggerPoint}</span>
            </div>
          {/if}

          <!-- 生成开销明细 -->
          {#if execution?.retryCount !== undefined || execution?.llmCost}
            <div class="log-row">
              <span class="log-label">生成开销:</span>
              <div class="cost-info">
                <span class="cost-item">
                  重试: <strong>{execution?.retryCount || 0} 次</strong>
                  {#if (execution?.retryCount || 0) > 0}
                    <span class="dim-text">(安全纠偏/重试)</span>
                  {/if}
                </span>
                {#if execution?.llmCost?.totalTokens}
                  {@const reasoningTokens = execution.llmCost.reasoningTokens || 0}
                  {@const compTokens = execution.llmCost.completionTokens || 0}
                  {@const contentTokens = execution.llmCost.contentTokens ?? (compTokens >= reasoningTokens ? compTokens - reasoningTokens : compTokens)}
                  <span class="cost-item">
                    Token: <strong>{execution.llmCost.totalTokens}</strong>
                    {#if reasoningTokens > 0}
                      <span class="dim-text">(生成={compTokens}: 思考={reasoningTokens} + 正文={contentTokens})</span>
                    {:else if contentTokens > 0}
                      <span class="dim-text">(正文: {contentTokens})</span>
                    {/if}
                  </span>
                  {#if execution.llmCost.promptTokens}
                    <span class="cost-item dim-text">提示: {execution.llmCost.promptTokens}</span>
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
              <span class="log-val status-green">未影响其他已打开工作簿</span>
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
              <span class="log-val status-green">COM 宏调用完成，临时模块已瞬时销毁清理</span>
            </div>
          {/if}

          <!-- 开放式 VBA 安全与回滚边界警示 -->
          <div class="boundary-warning-card">
            <div class="bw-header">
              <ShieldAlert size={14} color="#B25900" />
              <span class="bw-title">开放式 VBA 安全与回滚边界声明</span>
            </div>
            <p class="bw-body">
              工作簿快照仅保障目标工作簿本身的数据与格式原位回滚。若模型生成的开放式 VBA 包含外部文件读写、修改了其他工作簿、调用系统 API 或执行外部进程等副作用，快照无法自动撤回。宏的系统级行为仍需人工把关。
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

      <!-- 核心功能：多版本时间轴整本恢复控制栏 (自适应弹性排布) -->
      <div class="snapshot-footer">
        <div class="snap-info">
          <History size={14} color="#605E5C" />
          <span class="snap-label">版本回滚:</span>
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
            on:click={() => openRestoreConfirm(selectedSnapshotId || execution?.snapshot?.id || '')}
            disabled={isRestoring}
          >
            {#if restoreSuccess}
              <Check size={13} color="#107C41" />
              <span class="restore-success-text">已恢复整本！</span>
            {:else if isRestoring}
              <Loader2 size={13} class="spinner" /> 恢复中...
            {:else}
              <RotateCcw size={13} />
              <span>恢复到执行前</span>
            {/if}
          </button>
        {/if}
      </div>

      {#if showRestoreConfirm}
        <div class="rollback-confirm-panel">
          <div class="rcp-header">
            <ShieldAlert size={14} color="#C42B1C" />
            <span class="rcp-title">确认回滚整本工作簿？</span>
          </div>
          <p class="rcp-desc">
            确定将工作簿恢复到该宏执行前的快照吗？
            <br />
            ⚠️ <strong>注意</strong>：宏执行后产生的所有新增手工修改、单元格编辑都将被覆盖！
            <br />
            <span class="dim-text">（系统会在恢复前为当前状态保留一份紧急安全救援副本）</span>
          </p>
          <div class="rcp-actions">
            <button class="btn btn-sm" on:click={() => (showRestoreConfirm = false)}>取消</button>
            <button class="btn btn-sm btn-danger" on:click={confirmRestore}>确认恢复整本工作簿</button>
          </div>
        </div>
      {/if}

      {#if restoreErrorMsg}
        <div class="restore-error-bar">
          <AlertCircle size={14} color="#C42B1C" />
          <span>恢复失败: {restoreErrorMsg}</span>
        </div>
      {/if}

      <!-- 快照保护边界明确声明 -->
      <div class="snapshot-boundary-note">
        <span>ℹ️ 快照保护边界：快照完整备份和回滚目标工作簿本身的数据与结构；VBA 宏若涉及外部磁盘或网络操作，快照无法自动回滚。</span>
      </div>
    </div>
  {/if}
</div>

<!-- 双模态查看：独立全屏放大查看弹层 (用户明确要求的模态增强) -->
{#if isFullscreen}
  <div
    class="fullscreen-overlay"
    on:click|self={() => (isFullscreen = false)}
    on:keydown={(e) => e.key === 'Escape' && (isFullscreen = false)}
    role="presentation"
  >
    <div class="fullscreen-dialog" role="dialog" aria-modal="true" aria-label="执行详情与源码全屏查看">
      <div class="fs-header">
        <div class="fs-title-area">
          <span class="fs-title">执行详情与源码放大查看</span>
          <span class="badge {execution?.verificationStatus === 'verified' ? 'badge-green' : (execution?.error ? 'badge-red' : 'badge-amber')}">
            {execution?.verificationStatus === 'verified' ? '核验通过' : (execution?.error ? '执行异常' : '待确认')}
          </span>
          <span class="fs-subtitle">{execution?.targetWorkbookName || '当前工作簿'}</span>
        </div>

        <div class="fs-actions">
          <div class="fs-tab-group">
            <button
              class="fs-tab-btn {activeTab === 'code' ? 'active' : ''}"
              on:click={() => (activeTab = 'code')}
            >
              <Code size={13} />
              <span>VBA 源码</span>
            </button>
            <button
              class="fs-tab-btn {activeTab === 'readback' ? 'active' : ''}"
              on:click={() => (activeTab = 'readback')}
            >
              <SearchCheck size={13} />
              <span>写后核验</span>
            </button>
            <button
              class="fs-tab-btn {activeTab === 'audit' ? 'active' : ''}"
              on:click={() => (activeTab = 'audit')}
            >
              <FileText size={13} />
              <span>审计日志</span>
            </button>
          </div>

          <button class="btn btn-sm" on:click={handleCopyCurrent} title="复制当前展示代码">
            <Copy size={12} />
            <span>复制代码</span>
          </button>
          <button class="btn btn-sm" on:click={() => handleExportVba(codeViewMode)} title="导出为 .vba 文件">
            <Download size={12} />
            <span>导出.vba</span>
          </button>
          <button class="btn-icon" on:click={() => (isFullscreen = false)} title="还原/关闭全屏" aria-label="关闭">
            <Minimize2 size={16} />
          </button>
        </div>
      </div>

      <div class="fs-body">
        {#if activeTab === 'code'}
          <div class="fs-code-panel">
            <div class="fs-code-subbar">
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
              <div class="code-sub-actions">
                <button class="sub-action-btn" on:click={handleCopyOriginal}>复制原始</button>
                <button class="sub-action-btn" on:click={handleCopyExecuted}>复制执行</button>
              </div>
            </div>
            <pre class="fs-vba-pre"><code>{displayCode || '暂无可展示源码'}</code></pre>
          </div>
        {:else if activeTab === 'readback'}
          <div class="fs-readback-panel">
            {#if execution?.readback}
              <div class="fs-grid">
                <div class="fs-card">
                  <div class="fs-card-title">基本信息</div>
                  <div class="fs-card-row"><span>目标文件:</span> <strong>{execution.readback.targetWorkbookName}</strong></div>
                  <div class="fs-card-row"><span>目标表格:</span> <strong>{execution.readback.targetSheetName || '活动表'}</strong></div>
                  <div class="fs-card-row"><span>一致性:</span> <span class="pill {execution.readback.targetVerified ? 'pill-green' : 'pill-red'}">{execution.readback.targetVerified ? '核验通过' : '身份异常'}</span></div>
                </div>
                <div class="fs-card">
                  <div class="fs-card-title">写入区域读回</div>
                  <div class="fs-card-row"><span>使用区域:</span> <strong style="color: #107C41;">{execution.readback.usedRangeAddress || '未检测到'}</strong></div>
                  <div class="fs-card-row"><span>矩阵维度:</span> <span>{execution.readback.rowCount} 行 × {execution.readback.columnCount} 列</span></div>
                  <div class="fs-card-row"><span>起笔坐标:</span> <span>{execution.readback.startCell || 'N/A'}</span></div>
                </div>
                <div class="fs-card">
                  <div class="fs-card-title">特征分析</div>
                  <div class="fs-card-row"><span>公式计算:</span> <span>{execution.readback.hasFormulas ? '包含公式' : '纯静态值'}</span></div>
                  <div class="fs-card-row"><span>网格边框:</span> <span>{execution.readback.hasBorders ? '包含格式边框' : '无特殊边框'}</span></div>
                  <div class="fs-card-row"><span>单元格底色:</span> <span>{execution.readback.hasInteriorColor ? '包含背景填充色' : '默认无底色'}</span></div>
                </div>
              </div>
              {#if execution.readback.sampleValues && execution.readback.sampleValues.length > 0}
                <div class="fs-samples-section">
                  <div class="fs-card-title">读回单元格抽样值</div>
                  <div class="sample-tags">
                    {#each execution.readback.sampleValues as sample}
                      <code class="sample-code">{sample}</code>
                    {/each}
                  </div>
                </div>
              {/if}
              {#if execution.verificationNote}
                <div class="verification-box" style="margin-top: 14px;">
                  <span class="ver-label">核验综合结论:</span>
                  <span class="ver-text">{execution.verificationNote}</span>
                </div>
              {/if}
            {:else}
              <div class="dim-empty">尚未取得写后读回数据</div>
            {/if}
          </div>
        {:else if activeTab === 'audit'}
          <div class="fs-audit-panel">
            <div class="log-row">
              <span class="log-label">用户指令:</span>
              <span class="log-val"><strong>{prompt}</strong></span>
            </div>
            <div class="log-row">
              <span class="log-label">执行耗时:</span>
              <span class="log-val">{execution?.elapsedMs || 0} ms</span>
            </div>
            <div class="audit-box" style="margin: 12px 0;">
              <div class="audit-box-title">API 预算与模型审计:</div>
              <div class="audit-box-content">
                <div>• 最大输出 Token: <code>{execution?.apiAudit?.maxTokensStatus || '默认'}</code></div>
                <div>• 深度思考预算: <code>{execution?.apiAudit?.thinkingBudgetStatus || '默认'}</code></div>
                <div>• 发送历史消息: <code>{execution?.apiAudit?.historyCountSent ?? 'N/A'} 条</code></div>
              </div>
            </div>
            {#if execution?.error}
              <div class="error-block" style="margin-top: 12px;">
                <span class="log-label">完整错误堆栈:</span>
                <pre class="error-text">{execution.error}</pre>
              </div>
            {/if}
          </div>
        {/if}
      </div>
    </div>
  </div>
{/if}

<style>
  .execution-card {
    background: #ffffff;
    border-radius: var(--office-radius);
    border: 1px solid var(--office-border);
    margin: 6px 0;
    overflow: hidden;
    box-shadow: var(--office-shadow-sm);
    transition: all 0.15s ease;
  }

  .card-success {
    border-left: 3px solid var(--excel-green);
  }

  .card-unconfirmed {
    border-left: 3px solid var(--office-amber);
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
    transition: background 0.15s ease;
  }

  .summary-bar:hover {
    background: var(--office-card-subtle);
  }

  .status-left {
    display: flex;
    align-items: center;
    gap: 7px;
    min-width: 0;
    flex: 1;
  }

  .summary-text {
    font-size: var(--font-size-sm);
    font-weight: 500;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    line-height: 1.35;
  }

  .font-success {
    color: var(--excel-green);
  }

  .font-unconfirmed {
    color: var(--office-amber);
  }

  .font-error {
    color: var(--office-danger);
  }

  .font-running {
    color: var(--office-blue);
  }

  .summary-right-actions {
    display: flex;
    align-items: center;
    gap: 4px;
    flex-shrink: 0;
  }

  .btn-icon-subtle {
    padding: 4px;
    background: transparent;
    border: none;
    border-radius: var(--office-radius-xs);
    color: var(--office-muted);
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    transition: all 0.15s ease;
  }

  .btn-icon-subtle:hover {
    background: var(--office-hover);
    color: var(--office-text);
  }

  .expand-btn {
    display: flex;
    align-items: center;
    gap: 3px;
    font-size: var(--font-size-xs);
    color: var(--office-muted);
    background: none;
    border: none;
    cursor: pointer;
    padding: 3px 6px;
    border-radius: var(--office-radius-xs);
    flex-shrink: 0;
    font-weight: 500;
    transition: all 0.15s ease;
  }

  .expand-btn:hover {
    background: var(--office-hover);
    color: var(--office-text);
  }

  .expanded-panel {
    border-top: 1px solid var(--office-border);
    background: var(--office-card-subtle);
  }

  /* 自适应选项卡头部与响应式工具栏 */
  .panel-tabs {
    min-height: 34px;
    padding: 2px 10px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 1px solid var(--office-border);
    background: #f0f2f5;
    flex-wrap: wrap;
    gap: 4px;
  }

  .tabs-left {
    display: flex;
    height: 100%;
    gap: 2px;
  }

  .tab-btn {
    display: flex;
    align-items: center;
    gap: 4px;
    padding: 5px 8px;
    font-size: var(--font-size-xs);
    font-weight: 500;
    color: var(--office-muted);
    background: transparent;
    border: none;
    border-bottom: 2px solid transparent;
    cursor: pointer;
    transition: all 0.15s ease;
    border-radius: var(--office-radius-xs) var(--office-radius-xs) 0 0;
  }

  .tab-btn:hover {
    color: var(--office-text);
    background: rgba(0, 0, 0, 0.03);
  }

  .tab-btn.active {
    color: var(--excel-green);
    border-bottom-color: var(--excel-green);
    background: #ffffff;
    font-weight: 600;
  }

  .tabs-actions {
    display: flex;
    align-items: center;
    gap: 5px;
  }

  .btn-sm {
    padding: 2px 7px;
    font-size: var(--font-size-xs);
  }

  .copy-success-text {
    color: var(--excel-green);
    font-weight: 600;
  }

  /* 响应式自适应防截断：小于 440px 时按钮文字优雅缩略，纯图标+Tooltip，绝不换行错乱 */
  @media (max-width: 440px) {
    .btn-label-text {
      display: none;
    }
    .panel-tabs {
      padding: 2px 6px;
    }
    .tab-btn {
      padding: 5px 6px;
    }
  }

  /* 写后核验面板样式 */
  .readback-container {
    padding: 10px 12px;
    font-size: var(--font-size-xs);
    display: flex;
    flex-direction: column;
    gap: 6px;
    background: #ffffff;
  }

  .rb-row {
    display: flex;
    align-items: baseline;
    gap: 8px;
    overflow-wrap: anywhere;
  }

  .rb-label {
    color: var(--office-muted);
    min-width: 76px;
    flex-shrink: 0;
    font-weight: 500;
  }

  .rb-val {
    color: var(--office-text);
    word-break: break-all;
    display: flex;
    align-items: center;
    gap: 6px;
    flex-wrap: wrap;
  }

  .highlight-val {
    font-weight: 600;
    color: var(--excel-green);
  }

  .dim-text {
    font-weight: normal;
    color: var(--office-muted);
    font-size: var(--font-size-xs);
  }

  .pill {
    padding: 1px 6px;
    border-radius: var(--office-radius-xs);
    font-size: 10.5px;
    font-weight: 500;
    line-height: 1.35;
    white-space: nowrap;
  }

  .pill-green {
    background: var(--excel-light);
    color: var(--excel-dark);
    border: 1px solid var(--excel-light-border);
  }

  .pill-blue {
    background: var(--office-blue-light);
    color: var(--office-blue-dark);
    border: 1px solid var(--office-blue-border);
  }

  .pill-red {
    background: var(--office-danger-light);
    color: var(--office-danger);
    border: 1px solid var(--office-danger-border);
  }

  .pill-gray {
    background: #edebe9;
    color: var(--office-muted);
    border: 1px solid #e1dfdd;
  }

  .pill-amber {
    background: var(--office-amber-light);
    color: var(--office-amber);
    border: 1px solid var(--office-amber-border);
  }

  .sample-tags {
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
  }

  .sample-code {
    background: #f3f2f1;
    padding: 2px 6px;
    border-radius: var(--office-radius-xs);
    font-family: var(--font-family-code);
    font-size: 10px;
    color: #201f1e;
    border: 1px solid #e1dfdd;
  }

  .verification-box {
    margin-top: 4px;
    padding: 6px 9px;
    background: var(--office-amber-light);
    border: 1px solid var(--office-amber-border);
    border-radius: var(--office-radius-sm);
    display: flex;
    gap: 6px;
    font-size: var(--font-size-xs);
    line-height: var(--line-height-normal);
  }

  .ver-label {
    color: var(--office-amber);
    font-weight: 600;
    flex-shrink: 0;
  }

  .ver-text {
    color: #323130;
  }

  /* 代码区展示 */
  .code-view-wrapper {
    display: flex;
    flex-direction: column;
  }

  .code-sub-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #252526;
    padding: 5px 10px;
    border-bottom: 1px solid #333333;
    gap: 6px;
    flex-wrap: wrap;
  }

  .code-sub-tabs {
    display: flex;
    gap: 4px;
    align-items: center;
    flex-wrap: wrap;
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
    transition: all 0.15s ease;
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

  .code-sub-actions {
    display: flex;
    align-items: center;
    gap: 5px;
    flex-wrap: wrap;
  }

  .sub-action-btn {
    background: #2d2d30;
    border: 1px solid #3e3e42;
    color: #cccccc;
    font-size: 10px;
    padding: 2px 7px;
    border-radius: 3px;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 4px;
    transition: all 0.15s ease;
  }

  .sub-action-btn:hover {
    background: #3e3e42;
    color: #ffffff;
    border-color: #555555;
  }

  .sub-action-btn.copied {
    border-color: #4ec9b0;
    background: #1e3a34;
    color: #4ec9b0;
  }

  .action-text-copied {
    color: #4ec9b0;
    font-weight: 600;
  }

  .export-btn:hover {
    color: #4fc1ff;
    border-color: #007acc;
  }

  .hash-tag {
    font-family: var(--font-family-code);
    font-size: 9px;
    color: #4ec9b0;
    opacity: 0.9;
  }

  .code-container {
    position: relative;
    padding: 10px 12px;
    background: #1e1e1e;
    color: #d4d4d4;
    max-height: min(360px, 48vh);
    overflow-y: auto;
  }

  .floating-copy-btn {
    position: absolute;
    top: 8px;
    right: 12px;
    background: rgba(45, 45, 48, 0.9);
    backdrop-filter: blur(4px);
    border: 1px solid #3e3e42;
    color: #cccccc;
    font-size: 10px;
    padding: 3px 8px;
    border-radius: 4px;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 4px;
    transition: all 0.2s ease;
    z-index: 5;
  }

  .floating-copy-btn:hover {
    background: #3e3e42;
    color: #ffffff;
    border-color: #007acc;
  }

  .floating-copy-btn.copied {
    background: #1e3a34;
    border-color: #4ec9b0;
    color: #4ec9b0;
  }

  .vba-code {
    font-family: var(--font-family-code);
    font-size: var(--font-size-xs);
    line-height: var(--line-height-normal);
    white-space: pre-wrap;
    word-break: break-all;
    margin: 0;
  }

  /* 审计日志容器 */
  .log-container {
    padding: 10px 12px;
    font-size: var(--font-size-xs);
    display: flex;
    flex-direction: column;
    gap: 6px;
    max-height: min(320px, 42vh);
    overflow-y: auto;
  }

  .log-row {
    display: flex;
    gap: 8px;
    overflow-wrap: anywhere;
  }

  .flex-col {
    flex-direction: column;
    gap: 3px;
  }

  .log-label {
    color: var(--office-muted);
    min-width: 76px;
    font-weight: 500;
    flex-shrink: 0;
  }

  .log-val {
    color: var(--office-text);
    word-break: break-all;
  }

  .status-green {
    color: var(--excel-green);
    font-weight: 500;
  }

  .audit-box {
    background: #f8fafc;
    padding: 8px 10px;
    border-radius: var(--office-radius-sm);
    border: 1px solid #e2e8f0;
    margin: 4px 0;
  }

  .audit-box-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
  }

  .audit-box-title {
    font-weight: 600;
    color: #334155;
    font-size: var(--font-size-xs);
    margin-bottom: 2px;
  }

  .audit-box-content {
    font-size: 10.5px;
    line-height: 1.6;
    color: #475569;
  }

  .audit-constraints-list {
    font-size: 10.5px;
    line-height: 1.5;
    color: #475569;
    margin-top: 6px;
  }

  .constraint-item {
    margin-bottom: 4px;
  }

  .constraint-warning {
    color: var(--office-danger);
    margin-left: 12px;
    font-style: italic;
  }

  .error-chain-box {
    background: #fff8f8;
    padding: 8px 10px;
    border-radius: var(--office-radius-sm);
    border: 1px solid #fed7d7;
    margin: 4px 0;
  }

  .error-chain-label {
    color: var(--office-danger);
    font-weight: 600;
  }

  .error-chain-content {
    font-size: var(--font-size-xs);
    line-height: 1.55;
    color: #4a5568;
    margin-top: 2px;
  }

  .error-code {
    color: var(--office-danger);
    font-family: var(--font-family-code);
  }

  .partial-warning-box {
    background: #fffbe6;
    padding: 8px 10px;
    border-radius: var(--office-radius-sm);
    border: 1px solid #ffe58f;
    margin: 4px 0;
  }

  .partial-warning-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 6px;
  }

  .partial-warning-title {
    color: #d46b08;
    font-weight: 600;
    font-size: var(--font-size-xs);
  }

  .partial-warning-desc {
    font-size: var(--font-size-xs);
    color: #873800;
    margin-top: 4px;
    line-height: 1.4;
  }

  .trigger-point-val {
    color: var(--office-danger);
    font-weight: 500;
  }

  .hash-table {
    display: flex;
    flex-direction: column;
    gap: 4px;
    background: #f3f2f1;
    padding: 6px 8px;
    border-radius: var(--office-radius-xs);
    flex: 1;
  }

  .hash-item {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 10px;
    flex-wrap: wrap;
  }

  .hash-name {
    color: var(--office-muted);
    min-width: 55px;
  }

  .hash-item code {
    font-family: var(--font-family-code);
    font-size: 10px;
    color: #004e8c;
    word-break: break-all;
  }

  .boundary-warning-card {
    margin-top: 6px;
    padding: 8px 10px;
    background: #fff8f5;
    border: 1px solid #fed9cc;
    border-radius: var(--office-radius-sm);
  }

  .bw-header {
    display: flex;
    align-items: center;
    gap: 5px;
    margin-bottom: 3px;
  }

  .bw-title {
    font-size: var(--font-size-xs);
    font-weight: 600;
    color: var(--office-amber);
  }

  .bw-body {
    margin: 0;
    font-size: 10.5px;
    line-height: 1.45;
    color: var(--office-muted);
  }

  .step-list {
    margin: 2px 0 0 16px;
    padding: 0;
    color: #323130;
  }

  .error-block {
    flex-direction: column;
    gap: 4px;
  }

  .error-text {
    background: var(--office-danger-light);
    color: var(--office-danger);
    padding: 6px 8px;
    border-radius: var(--office-radius-xs);
    font-family: var(--font-family-code);
    font-size: var(--font-size-xs);
    white-space: pre-wrap;
    border: 1px solid var(--office-danger-border);
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
    font-size: var(--font-size-xs);
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
    font-size: var(--font-size-xs);
    border: 1px solid #c2e2cc;
    border-radius: var(--office-radius-xs);
    outline: none;
    background: #ffffff;
  }

  .desc-input {
    height: 24px;
    padding: 0 8px;
    font-size: var(--font-size-xs);
    border: 1px solid #e1e1e1;
    border-radius: var(--office-radius-xs);
    background: #ffffff;
  }

  /* 版本回滚底部控制栏 */
  .snapshot-footer {
    padding: 8px 12px;
    background: #ffffff;
    border-top: 1px solid var(--office-border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    flex-wrap: wrap;
  }

  .snap-info {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: var(--font-size-xs);
    color: var(--office-muted);
    min-width: 0;
    flex: 1;
    flex-wrap: wrap;
  }

  .snap-select {
    height: 24px;
    font-size: var(--font-size-xs);
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-xs);
    padding: 0 4px;
    background: #ffffff;
    max-width: min(200px, 100%);
    color: var(--office-text);
  }

  .snap-tag {
    font-weight: 500;
    color: var(--office-text);
  }

  .snap-none {
    color: var(--office-dim);
    font-style: italic;
  }

  .btn-rollback {
    color: var(--office-amber);
    background: var(--office-amber-light);
    border-color: var(--office-amber-border);
    font-weight: 600;
    flex-shrink: 0;
  }

  .btn-rollback:hover:not(:disabled) {
    background: #ffeedb;
  }

  .restore-success-text {
    color: var(--excel-green);
    font-weight: 600;
  }

  /* 部分修改常驻横幅样式 */
  .partial-mod-banner {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    padding: 8px 12px;
    background: #fffbe6;
    border-top: 1px solid #ffe58f;
    border-bottom: 1px solid #ffe58f;
    flex-wrap: wrap;
  }

  .partial-mod-left {
    display: flex;
    align-items: flex-start;
    gap: 8px;
    flex: 1;
    min-width: 200px;
  }

  .partial-mod-text {
    display: flex;
    flex-direction: column;
    gap: 2px;
  }

  .partial-mod-title {
    font-size: var(--font-size-xs);
    font-weight: 600;
    color: #d46b08;
  }

  .partial-mod-desc {
    font-size: var(--font-size-xs);
    color: #873800;
    line-height: 1.4;
  }

  .partial-mod-right {
    flex-shrink: 0;
  }

  .dim-empty {
    color: var(--office-muted);
    font-style: italic;
    padding: 4px 0;
  }

  :global(.spinner) {
    animation: spin 1s linear infinite;
  }

  .rollback-confirm-panel {
    background: #fff8f5;
    border: 1px solid #f8d0c0;
    border-radius: var(--office-radius-sm);
    padding: 10px 12px;
    margin: 8px 12px;
  }

  .rcp-header {
    display: flex;
    align-items: center;
    gap: 6px;
    margin-bottom: 4px;
  }

  .rcp-title {
    font-size: var(--font-size-sm);
    font-weight: 600;
    color: var(--office-danger);
  }

  .rcp-desc {
    font-size: var(--font-size-xs);
    color: var(--office-muted);
    line-height: var(--line-height-normal);
    margin: 0 0 8px 0;
  }

  .rcp-actions {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
  }

  .restore-error-bar {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 6px 12px;
    background: var(--office-danger-light);
    color: var(--office-danger);
    font-size: var(--font-size-xs);
    margin: 4px 12px;
    border-radius: var(--office-radius-xs);
    border: 1px solid var(--office-danger-border);
  }

  @keyframes spin {
    from { transform: rotate(0deg); }
    to { transform: rotate(360deg); }
  }

  .snapshot-boundary-note {
    font-size: 10.5px;
    color: var(--office-muted);
    padding: 6px 12px;
    background: #fbfbfb;
    border-top: 1px dashed var(--office-border);
    line-height: 1.4;
  }

  /* 双模态查看：独立全屏放大遮罩与弹层 */
  .fullscreen-overlay {
    position: fixed;
    top: 0;
    left: 0;
    width: 100vw;
    height: 100vh;
    background: rgba(0, 0, 0, 0.45);
    backdrop-filter: blur(3px);
    z-index: 1000;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 16px;
  }

  .fullscreen-dialog {
    width: min(94vw, 900px);
    height: min(92vh, 800px);
    background: #ffffff;
    border-radius: var(--office-radius-lg);
    box-shadow: var(--office-shadow-lg);
    display: flex;
    flex-direction: column;
    overflow: hidden;
    border: 1px solid var(--office-border);
    animation: modalPop 0.18s ease-out;
  }

  .fs-header {
    height: 48px;
    padding: 0 16px;
    border-bottom: 1px solid var(--office-border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #ffffff;
    flex-shrink: 0;
    gap: 12px;
  }

  .fs-title-area {
    display: flex;
    align-items: center;
    gap: 8px;
    min-width: 0;
  }

  .fs-title {
    font-size: var(--font-size-md);
    font-weight: 600;
    color: var(--office-text);
  }

  .fs-subtitle {
    font-size: var(--font-size-xs);
    color: var(--office-muted);
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .fs-actions {
    display: flex;
    align-items: center;
    gap: 8px;
    flex-shrink: 0;
  }

  .fs-tab-group {
    display: inline-flex;
    background: #f0f2f5;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-sm);
    padding: 2px;
    gap: 2px;
  }

  .fs-tab-btn {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    padding: 3px 8px;
    font-size: var(--font-size-xs);
    border: none;
    border-radius: 3px;
    background: transparent;
    color: var(--office-muted);
    cursor: pointer;
    transition: all 0.15s ease;
  }

  .fs-tab-btn.active {
    background: #ffffff;
    color: var(--excel-green);
    font-weight: 600;
    box-shadow: var(--office-shadow-sm);
  }

  .fs-body {
    flex: 1;
    overflow-y: auto;
    display: flex;
    flex-direction: column;
    background: #fcfcfc;
  }

  .fs-code-panel {
    display: flex;
    flex-direction: column;
    height: 100%;
  }

  .fs-code-subbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 8px 16px;
    background: #252526;
    border-bottom: 1px solid #333;
    gap: 8px;
  }

  .fs-vba-pre {
    flex: 1;
    margin: 0;
    padding: 16px 20px;
    background: #1e1e1e;
    color: #d4d4d4;
    font-family: var(--font-family-code);
    font-size: var(--font-size-sm);
    line-height: 1.55;
    overflow: auto;
    white-space: pre-wrap;
    word-break: break-all;
  }

  .fs-readback-panel {
    padding: 20px;
    display: flex;
    flex-direction: column;
    gap: 16px;
  }

  .fs-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 12px;
  }

  .fs-card {
    background: #ffffff;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    padding: 14px;
    display: flex;
    flex-direction: column;
    gap: 8px;
    box-shadow: var(--office-shadow-sm);
  }

  .fs-card-title {
    font-size: var(--font-size-sm);
    font-weight: 600;
    color: var(--office-text);
    border-bottom: 1px solid var(--office-border-subtle);
    padding-bottom: 6px;
  }

  .fs-card-row {
    display: flex;
    justify-content: space-between;
    font-size: var(--font-size-xs);
    color: var(--office-muted);
  }

  .fs-card-row strong {
    color: var(--office-text);
  }

  .fs-samples-section {
    background: #ffffff;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    padding: 14px;
  }

  .fs-audit-panel {
    padding: 20px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: var(--font-size-sm);
  }
</style>
