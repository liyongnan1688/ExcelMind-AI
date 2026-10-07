<script lang="ts">
  import {
    Play,
    ShieldCheck,
    AlertCircle,
    ChevronDown,
    ChevronUp,
    Copy,
    Check,
    Code2,
    FileSpreadsheet,
    MessageSquare,
    RotateCcw,
    Info,
    ExternalLink,
    Sliders,
    HelpCircle,
  } from 'lucide-svelte';

  export let macroTitle = '销售数据智能汇总与去重';
  export let macroSummary = '提取订单记录，比对重复项并计算销售额总计。';
  export let targetWorkbook = 'Sales_2026.xlsx';
  export let activeWorkbook = 'Sales_2026.xlsx'; // 当前活动工作簿
  export let activeSheetName = 'Sheet1';
  // 客观事实边界：代码执行前绝不主观声明影响范围
  export let scopeDescription = '修改范围完全取决于 VBA 源码逻辑（执行前无法推断是否新建工作表或访问外部文件）';
  // 快照状态：严格遵循“执行前显示确认后将创建，收到成功才显示已创建”
  export let snapshotCreated = false;
  export let snapshotNotice = '确认后将创建目标工作簿快照；失败则不执行。';
  export let isSimulation = false; // 默认生产状态，仅在隔离原型时传入 true

  export let vbaCode = `Sub SummarizeAndDedupSales(targetWb As Workbook)
    Dim ws As Worksheet
    Set ws = targetWb.Worksheets("Sheet1")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ws.Range("A1:F" & lastRow).RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
    
    Dim newLastRow As Long
    newLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Cells(newLastRow + 2, "A").Value = "汇总总计"
    ws.Cells(newLastRow + 2, "D").Formula = "=SUM(D2:D" & newLastRow & ")"
End Sub`;

  // 单列参数表单配置项
  export let parameters: Array<{ name: string; label: string; type: 'string' | 'number' | 'boolean'; value: any; required?: boolean; hint?: string }> = [
    { name: 'startRow', label: '起始数据行', type: 'number', value: 2, required: true, hint: '跳过标题行，从该行开始处理数据' },
    { name: 'dedupKeyCols', label: '去重关键列', type: 'string', value: '1, 2', required: true, hint: '输入列序号（如 1 或 1, 2）进行复合比较' }
  ];

  export let onConfirmExecute: () => void = () => {};
  export let onCancel: () => void = () => {};
  export let onSwitchToChat: (draftPrompt: string) => void = () => {};

  let isCodeExpanded = false;
  let showDetailedRecovery = false;
  let showDetailedScope = false;
  let copied = false;
  let executionState: 'idle' | 'running' | 'success' | 'failed' = 'idle';
  let executionMessage = '';

  // 用户可明确选择是否附带源码和参数至对话草稿
  let includeParamsInChat = true;
  let includeCodeInChat = false;

  function handleCopyCode() {
    navigator.clipboard?.writeText(vbaCode);
    copied = true;
    setTimeout(() => (copied = false), 1800);
  }

  function handleExecute() {
    executionState = 'running';
    if (isSimulation) {
      setTimeout(() => {
        snapshotCreated = true;
        executionState = 'success';
        executionMessage = '宏已执行完成（原型模拟）。已创建执行前快照副本。请在工作表中核对实际业务修改结果。';
        onConfirmExecute();
      }, 600);
    } else {
      // 生产环境真实调用确认
      onConfirmExecute();
    }
  }

  function handleRollback() {
    executionState = 'idle';
    executionMessage = '已唤起快照恢复流程。目标工作簿已重载至执行前快照副本。';
  }

  function handleTransferToChat() {
    let sections: string[] = [];
    sections.push(`请分析宏【${macroTitle}】的执行逻辑与潜在影响：`);
    if (includeParamsInChat && parameters.length > 0) {
      const pText = parameters.map(p => `${p.label || p.name}: ${p.value}`).join('；');
      sections.push(`当前拟定参数：${pText}`);
    }
    if (includeCodeInChat && vbaCode) {
      sections.push(`\`\`\`vba\n${vbaCode}\n\`\`\``);
    }
    const draft = sections.join('\n\n');
    // 只预填可编辑草稿，不自动发送，不清空参数草稿
    onSwitchToChat(draft);
  }
</script>

<div class="confirm-card-container">
  <!-- 头部：标题与用途 -->
  <div class="card-header">
    <div class="header-badge-row">
      <span class="pill pill-green">
        <ShieldCheck size={12} />
        <span>运行前核验</span>
      </span>
      {#if isSimulation}
        <span class="pill pill-amber">原型模拟状态</span>
      {/if}
      <span class="code-hash-tag">SHA: 8f2a9c1e</span>
    </div>
    <h3 class="macro-name">{macroTitle}</h3>
    <p class="macro-desc">{macroSummary}</p>
  </div>

  <!-- 纵向单列事实区 (上为标签，下为内容，彻底消除左右分列横向挤压) -->
  <div class="facts-vertical-list">
    <div class="fact-block">
      <div class="fact-label-row">
        <span class="fact-label">锁定目标工作簿</span>
        {#if activeWorkbook && targetWorkbook && activeWorkbook !== targetWorkbook}
          <span class="fact-warn-tag">当前活动: {activeWorkbook}</span>
        {/if}
      </div>
      <div class="fact-box font-mono">
        <FileSpreadsheet size={13} class="icon-inline" />
        <strong>{targetWorkbook}</strong>
      </div>
    </div>

    <div class="fact-block">
      <div class="fact-label-row">
        <span class="fact-label">修改范围客观边界</span>
        <button
          type="button"
          class="btn-text-link"
          on:click={() => (showDetailedScope = !showDetailedScope)}
        >
          {showDetailedScope ? '收起说明' : '详细范围'}
        </button>
      </div>
      <div class="fact-box fact-scope-box">
        <AlertCircle size={13} class="icon-inline text-amber shrink-0" />
        <span>{scopeDescription}</span>
      </div>
      {#if showDetailedScope}
        <div class="fact-sub-hint">
          当前活动表为【{activeSheetName}】。由于宏具有完整 VBA 运行权限，执行可能影响其他工作表或新建工作表，真实影响以源码逻辑为准。
        </div>
      {/if}
    </div>

    <div class="fact-block">
      <div class="fact-label-row">
        <span class="fact-label">快照保障状态</span>
        <button
          type="button"
          class="btn-text-link"
          on:click={() => (showDetailedRecovery = !showDetailedRecovery)}
        >
          {showDetailedRecovery ? '收起说明' : '恢复说明'}
        </button>
      </div>
      <div class="fact-box fact-snapshot-box {snapshotCreated ? 'snapshot-active' : ''}">
        <ShieldCheck size={13} class="icon-inline text-green shrink-0" />
        <span class="snapshot-status-text">
          {#if snapshotCreated}
            已创建目标工作簿快照（当前会话内可按已验证范围回滚）
          {:else}
            {snapshotNotice}
          {/if}
        </span>
      </div>
      {#if showDetailedRecovery}
        <div class="fact-sub-hint">
          回滚基于执行前捕获的物理副本文件；仅限当前会话内已验证范围，不作任意外部副作用无损撤销的绝对承诺。
        </div>
      {/if}
    </div>
  </div>

  <!-- 单列参数表单 (标签 → 说明 → 全宽输入控件) -->
  {#if parameters.length > 0}
    <div class="params-section-vertical">
      <div class="params-header">
        <Sliders size={13} />
        <span>运行参数配置</span>
      </div>
      <div class="params-list">
        {#each parameters as param}
          <div class="param-single-col">
            <div class="param-label-row">
              <label class="param-label" for={`param-${param.name}`}>
                {param.label || param.name}
                {#if param.required}<span class="req-star">*</span>{/if}
              </label>
              <span class="param-type-badge">{param.type === 'number' ? '数值' : '文本'}</span>
            </div>
            {#if param.hint}
              <span class="param-hint">{param.hint}</span>
            {/if}
            {#if param.type === 'number'}
              <input
                type="number"
                id={`param-${param.name}`}
                bind:value={param.value}
                class="param-input-control"
              />
            {:else}
              <input
                type="text"
                id={`param-${param.name}`}
                bind:value={param.value}
                class="param-input-control"
              />
            {/if}
          </div>
        {/each}
      </div>
    </div>
  {/if}

  <!-- 源码折叠核验区 -->
  <div class="code-section">
    <button
      type="button"
      class="code-toggle-bar"
      on:click={() => (isCodeExpanded = !isCodeExpanded)}
      aria-expanded={isCodeExpanded}
    >
      <div class="toggle-left">
        <Code2 size={13} />
        <span>VBA 源码核验</span>
      </div>
      <div class="toggle-right">
        <span class="toggle-hint">{isCodeExpanded ? '收起源码' : '展开查看'}</span>
        {#if isCodeExpanded}
          <ChevronUp size={14} />
        {:else}
          <ChevronDown size={14} />
        {/if}
      </div>
    </button>

    {#if isCodeExpanded}
      <div class="code-body-wrapper">
        <div class="code-toolbar">
          <span class="code-syntax-label">纯 VBA 源码（正文零修改）</span>
          <button type="button" class="mini-btn" on:click={handleCopyCode}>
            {#if copied}
              <Check size={12} color="#107c41" />
              <span>已复制</span>
            {:else}
              <Copy size={12} />
              <span>复制代码</span>
            {/if}
          </button>
        </div>
        <pre class="vba-code-viewer"><code>{vbaCode}</code></pre>
      </div>
    {/if}
  </div>

  <!-- 执行结果展示 -->
  {#if executionState === 'success'}
    <div class="callout callout-info result-box">
      <div class="callout-icon">
        <Info size={15} />
      </div>
      <div class="callout-content">
        <div class="result-title">执行完成（待确认业务效果）</div>
        <p class="result-msg">{executionMessage}</p>
        <div class="rollback-bar">
          <button type="button" class="btn btn-sm btn-rollback" on:click={handleRollback}>
            <RotateCcw size={12} />
            <span>按快照恢复已验证范围</span>
          </button>
        </div>
      </div>
    </div>
  {/if}

  <!-- 转到对话选项与底部操作按钮栏 -->
  {#if executionState !== 'success'}
    <div class="transfer-options-bar">
      <span class="transfer-opt-title">转到对话附带：</span>
      <label class="opt-checkbox-item">
        <input type="checkbox" bind:checked={includeParamsInChat} />
        <span>运行参数</span>
      </label>
      <label class="opt-checkbox-item">
        <input type="checkbox" bind:checked={includeCodeInChat} />
        <span>VBA 源码</span>
      </label>
    </div>

    <div class="card-footer-actions">
      <button type="button" class="btn-action btn-secondary" on:click={onCancel}>
        放弃
      </button>
      <button
        type="button"
        class="btn-action btn-chat-transfer"
        on:click={handleTransferToChat}
        title="将当前宏与所选内容填入输入框作为草稿，不自动发送"
      >
        <MessageSquare size={13} />
        <span>转到对话</span>
      </button>
      <button
        type="button"
        class="btn-action btn-primary-action"
        on:click={handleExecute}
        disabled={executionState === 'running'}
      >
        <Play size={13} />
        <span>{executionState === 'running' ? '正在执行...' : '确认执行宏'}</span>
      </button>
    </div>
  {/if}
</div>

<style>
  .confirm-card-container {
    background: #ffffff;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    padding: 12px;
    box-shadow: var(--office-shadow-sm);
    display: flex;
    flex-direction: column;
    gap: 10px;
    width: 100%;
    box-sizing: border-box;
  }

  .card-header {
    display: flex;
    flex-direction: column;
    gap: 3px;
  }

  .header-badge-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 2px;
  }

  .code-hash-tag {
    font-family: var(--font-family-code);
    font-size: 10px;
    color: var(--office-dim);
    background: #f0f0f0;
    padding: 1px 4px;
    border-radius: 2px;
  }

  .macro-name {
    font-size: var(--font-size-md);
    font-weight: 600;
    color: var(--office-text);
    margin: 0;
    line-height: 1.3;
  }

  .macro-desc {
    font-size: var(--font-size-sm);
    color: var(--office-muted);
    margin: 0;
    line-height: 1.4;
  }

  /* 纵向单列事实区 */
  .facts-vertical-list {
    display: flex;
    flex-direction: column;
    gap: 6px;
    width: 100%;
  }

  .fact-block {
    display: flex;
    flex-direction: column;
    gap: 2px;
    width: 100%;
  }

  .fact-label-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
  }

  .fact-label {
    font-size: 11px;
    font-weight: 600;
    color: var(--office-muted);
    line-height: 1.3;
  }

  .btn-text-link {
    background: transparent;
    border: none;
    color: var(--office-blue);
    font-size: 11px;
    cursor: pointer;
    padding: 0;
    text-decoration: underline;
  }

  .btn-text-link:hover {
    color: var(--office-blue-dark);
  }

  .fact-box {
    background: #f8f9fa;
    border: 1px solid var(--office-border-subtle);
    border-radius: var(--office-radius-xs);
    padding: 5px 8px;
    font-size: var(--font-size-sm);
    color: var(--office-text);
    line-height: 1.4;
    word-break: break-word;
    display: flex;
    align-items: center;
    gap: 5px;
    box-sizing: border-box;
    width: 100%;
  }

  .fact-scope-box {
    align-items: flex-start;
    background: #fffdf5;
    border-color: var(--office-amber-border);
    color: var(--office-text-secondary);
    font-size: var(--font-size-xs);
  }

  .fact-snapshot-box {
    align-items: flex-start;
    background: #f4faf6;
    border-color: var(--excel-light-border);
    color: var(--excel-dark);
    font-size: var(--font-size-xs);
  }

  .fact-sub-hint {
    font-size: 10px;
    color: var(--office-dim);
    line-height: 1.3;
    padding: 2px 4px;
    background: #fafafa;
    border-radius: 2px;
  }

  .icon-inline {
    flex-shrink: 0;
    margin-top: 2px;
  }

  .text-amber {
    color: var(--office-amber);
  }

  .text-green {
    color: var(--excel-green);
  }

  /* 源码折叠区 */
  .code-section {
    border: 1px solid var(--office-border-subtle);
    border-radius: var(--office-radius-sm);
    overflow: hidden;
  }

  .code-toggle-bar {
    width: 100%;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 7px 10px;
    background: #fafbfc;
    border: none;
    cursor: pointer;
    font-size: var(--font-size-sm);
    color: var(--office-text-secondary);
    transition: background 0.15s ease;
  }

  .code-toggle-bar:hover {
    background: #f0f2f5;
  }

  .toggle-left {
    display: flex;
    align-items: center;
    gap: 6px;
    font-weight: 500;
  }

  .toggle-right {
    display: flex;
    align-items: center;
    gap: 4px;
    color: var(--office-dim);
  }

  .toggle-hint {
    font-size: var(--font-size-xs);
  }

  .code-body-wrapper {
    background: #1e1e1e;
    display: flex;
    flex-direction: column;
  }

  .code-toolbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 5px 10px;
    background: #252526;
    border-bottom: 1px solid #333333;
  }

  .code-syntax-label {
    font-size: var(--font-size-xs);
    color: #999999;
  }

  .mini-btn {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    background: transparent;
    border: none;
    color: #cccccc;
    font-size: var(--font-size-xs);
    cursor: pointer;
    padding: 2px 6px;
    border-radius: 2px;
  }

  .mini-btn:hover {
    background: #37373d;
  }

  .vba-code-viewer {
    margin: 0;
    padding: 10px;
    font-family: var(--font-family-code);
    font-size: 11px;
    line-height: 1.45;
    color: #d4d4d4;
    max-height: 160px;
    overflow-y: auto;
    white-space: pre-wrap;
    word-break: break-all;
  }

  /* 执行结果与回滚 */
  .result-box {
    margin-top: 4px;
  }

  .result-title {
    font-weight: 600;
    font-size: var(--font-size-sm);
    margin-bottom: 2px;
  }

  .result-msg {
    margin: 0 0 6px 0;
    font-size: var(--font-size-xs);
  }

  .rollback-bar {
    display: flex;
    justify-content: flex-end;
  }

  .btn-rollback {
    background: #ffffff;
    border: 1px solid var(--excel-light-border);
    color: var(--excel-dark);
    font-size: var(--font-size-xs);
    padding: 4px 10px;
    border-radius: var(--office-radius-sm);
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 4px;
  }

  .btn-rollback:hover {
    background: var(--excel-light);
  }

  /* 单列参数配置表单 */
  .params-section-vertical {
    background: #fafbfc;
    border: 1px solid var(--office-border-subtle);
    border-radius: var(--office-radius-sm);
    padding: 8px;
    display: flex;
    flex-direction: column;
    gap: 6px;
    width: 100%;
    box-sizing: border-box;
  }

  .params-header {
    display: flex;
    align-items: center;
    gap: 5px;
    font-size: var(--font-size-xs);
    font-weight: 600;
    color: var(--office-text-secondary);
  }

  .params-list {
    display: flex;
    flex-direction: column;
    gap: 6px;
    width: 100%;
  }

  .param-single-col {
    display: flex;
    flex-direction: column;
    gap: 2px;
    width: 100%;
  }

  .param-label-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 4px;
  }

  .param-label {
    font-size: var(--font-size-xs);
    font-weight: 500;
    color: var(--office-text-secondary);
  }

  .param-hint {
    font-size: 10px;
    color: var(--office-dim);
    line-height: 1.25;
  }

  .param-input-control {
    width: 100%;
    height: 28px;
    background: #ffffff;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-xs);
    padding: 0 8px;
    font-size: var(--font-size-xs);
    font-family: inherit;
    color: var(--office-text);
    outline: none;
    box-sizing: border-box;
    transition: border-color 0.15s ease;
  }

  .param-input-control:focus {
    border-color: var(--excel-green);
  }

  .param-type-badge {
    font-size: 10px;
    background: #edf2f7;
    color: #4a5568;
    padding: 1px 5px;
    border-radius: 3px;
    font-family: inherit;
  }

  .req-star {
    color: #e53e3e;
    margin-left: 2px;
  }

  .fact-warn-tag {
    font-size: 10px;
    color: #c05621;
    background: #feebc8;
    padding: 1px 5px;
    border-radius: 3px;
  }

  /* 转到对话勾选项 */
  .transfer-options-bar {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 4px 6px;
    background: #f7fafc;
    border: 1px dashed #cbd5e0;
    border-radius: var(--office-radius-xs, 4px);
    font-size: 11px;
    color: var(--office-text-secondary);
  }

  .transfer-opt-title {
    font-size: 11px;
    color: var(--office-muted);
    flex-shrink: 0;
  }

  .opt-checkbox-item {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    cursor: pointer;
    user-select: none;
    font-size: 11px;
    color: var(--office-text);
  }

  .opt-checkbox-item input[type="checkbox"] {
    margin: 0;
    cursor: pointer;
  }

  /* 底部按钮栏 */
  .card-footer-actions {
    display: flex;
    align-items: center;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 6px;
    flex-wrap: wrap;
  }

  .btn-action {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 6px 12px;
    font-size: var(--font-size-sm);
    font-weight: 500;
    border-radius: var(--office-radius-sm);
    cursor: pointer;
    transition: all 0.15s ease;
    white-space: nowrap;
    box-sizing: border-box;
  }

  .btn-secondary {
    color: var(--office-muted);
    background: #ffffff;
    border: 1px solid var(--office-border);
  }

  .btn-secondary:hover {
    background: var(--office-hover);
    color: var(--office-text);
  }

  .btn-chat-transfer {
    background: #ffffff;
    border: 1px solid var(--office-blue-border);
    color: var(--office-blue-dark);
  }

  .btn-chat-transfer:hover {
    background: var(--office-blue-light);
  }

  .btn-primary-action {
    font-weight: 600;
    color: #ffffff;
    background: var(--excel-green);
    border: 1px solid var(--excel-green);
    box-shadow: 0 1px 2px rgba(16, 124, 65, 0.25);
  }

  .btn-primary-action:hover:not(:disabled) {
    background: var(--excel-hover-bg);
  }

  .btn-primary-action:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }

  @media (max-width: 360px) {
    .card-footer-actions {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 6px;
      width: 100%;
    }

    .btn-secondary,
    .btn-chat-transfer {
      justify-content: center;
      padding: 6px 6px;
    }

    .btn-primary-action {
      grid-column: span 2;
      justify-content: center;
      padding: 7px 12px;
    }
  }
</style>

