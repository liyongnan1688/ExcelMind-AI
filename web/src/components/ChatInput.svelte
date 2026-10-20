<script lang="ts">
  import { Zap, MessageSquare, Code, Plus, Info } from 'lucide-svelte';
  import {
    bridge,
    type SelectionContextData,
    type SelectionSendOptions,
    type SelectionSnapshot,
    type FormulaAuditResult,
    type WorkbookOverviewResult,
  } from '../services/bridge';
  import { formatSelectionContextForPrompt } from '../services/llm';
  import type { MacroReferenceData } from '../services/conversationManager';

  export let disabled = false;
  export let workbook: any = null;
  export let referencedMacro: MacroReferenceData | null = null;
  export let onSend: (
    text: string,
    mode?: 'AUTOMATION' | 'CHAT',
    selectionSnapshot?: SelectionSnapshot,
    macroReference?: MacroReferenceData
  ) => Promise<boolean | void> | boolean | void;

  $: isUnsavedWorkbook = Boolean(workbook && (!workbook.isSaved || !workbook.fullName || !workbook.fullName.includes('\\')));

  let inputText = '';
  let currentMode: 'AUTOMATION' | 'CHAT' = 'AUTOMATION';
  let showMacroCodeModal = false;
  let showPlusMenu = false;
  let plusContainerEl: HTMLElement | null = null;

  // 选区上下文只读状态
  let selectionContext: SelectionContextData | null = null;
  let isLoadingSelection = false;
  let selectionError = '';
  let showDetails = false;

  // R1b: 刷新原区域状态与并发防护序列号
  let isRefreshingArea = false;
  let refreshError = '';
  let refreshSeq = 0;

  let sendOptions: SelectionSendOptions = {
    includeStructure: true,
    includeSamples: false, // 默认不勾选样本
    includeFormulas: false, // 默认不勾选公式
    firstRowAsHeader: true, // 默认首行为候选表头
  };

  // Phase 2: 公式血缘体检与全簿架构透视状态
  let isAuditingFormula = false;
  let isScanningOverview = false;
  let auditModalData: FormulaAuditResult | null = null;
  let showAuditModal = false;

  async function handleAuditCurrentFormula() {
    if (!selectionContext || isAuditingFormula) return;
    isAuditingFormula = true;
    selectionError = '';
    try {
      const cell = selectionContext.address.split(':')[0];
      const res = await bridge.traceDependencies({
        targetWorkbookName: selectionContext.workbookName,
        targetWorkbookFullName: selectionContext.workbookFullName,
        sheetName: selectionContext.sheetName,
        cellAddress: cell,
      });
      if (res) {
        auditModalData = res;
        showAuditModal = true;
      } else {
        selectionError = '未检测到公式或解析失败';
      }
    } catch (e: any) {
      selectionError = '公式血缘分析异常: ' + (e?.message || e);
    } finally {
      isAuditingFormula = false;
    }
  }

  async function handleScanWorkbookOverview() {
    if (isScanningOverview || disabled) return;
    isScanningOverview = true;
    selectionError = '';
    try {
      const res = await bridge.getWorkbookOverview({
        targetWorkbookName: workbook?.name,
        targetWorkbookFullName: workbook?.fullName,
      });
      if (res && res.blueprintSummary) {
        inputText = `【工作簿全局架构透视】\n${res.blueprintSummary}\n\n请根据以上各工作表结构，帮我：`;
        currentMode = 'CHAT';
        if (textareaEl) {
          textareaEl.focus();
        }
      } else {
        selectionError = '未能读取到工作簿结构，请确保已打开工作簿';
      }
    } catch (e: any) {
      selectionError = '获取工作簿透视失败: ' + (e?.message || e);
    } finally {
      isScanningOverview = false;
    }
  }

  function handleInsertExplanationPrompt() {
    if (!auditModalData) return;
    inputText = auditModalData.explanationPrompt;
    currentMode = 'CHAT';
    showAuditModal = false;
    if (textareaEl) {
      textareaEl.focus();
    }
  }

  function formatCaptureTime(ts?: number): string {
    if (!ts) return '';
    const d = new Date(ts);
    const pad = (n: number) => (n < 10 ? '0' + n : n);
    return `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
  }

  // 【附加当前选区】（替换动作：读取当前 Excel 活动选区）
  async function handleAttachSelection() {
    if (disabled || isLoadingSelection || isRefreshingArea) return;
    isLoadingSelection = true;
    selectionError = '';
    refreshError = '';
    const currentSeq = ++refreshSeq;

    try {
      const res = await bridge.getSelectionContext(5, 15);
      if (res.ok && res.data) {
        if (currentSeq >= refreshSeq) {
          selectionContext = res.data;
          sendOptions = {
            includeStructure: true,
            includeSamples: false,
            includeFormulas: false,
            firstRowAsHeader: true,
          };
          refreshError = '';
        }
      } else {
        if (currentSeq >= refreshSeq) {
          selectionError = res.error || '获取选区失败';
        }
      }
    } catch (err: any) {
      if (currentSeq >= refreshSeq) {
        selectionError = err.message || '获取选区异常';
      }
    } finally {
      if (currentSeq >= refreshSeq) {
        isLoadingSelection = false;
      }
    }
  }

  // 【刷新原区域】（定向动作：读取该附件保存的原工作簿、原工作表和原区域地址）
  async function handleRefreshOriginalArea() {
    if (disabled || isRefreshingArea || isLoadingSelection || !selectionContext) return;
    const currentAttId = selectionContext.attachmentId;
    const currentSeq = ++refreshSeq;

    isRefreshingArea = true;
    refreshError = '';

    try {
      const res = await bridge.getSelectionContext({
        sampleRows: selectionContext.sampleRowCount || 5,
        sampleCols: selectionContext.sampleColumnCount || 15,
        targetWorkbookFullName: selectionContext.workbookFullName,
        targetWorkbookName: selectionContext.workbookName,
        targetSheetName: selectionContext.sheetName,
        targetAddress: selectionContext.address,
        attachmentId: currentAttId,
      });

      // 1. 若当前卡片已被移除或被新选区替换，丢弃过时结果，绝不重新挂回已移除卡片
      if (!selectionContext || selectionContext.attachmentId !== currentAttId) {
        return;
      }

      // 2. 若序列号小于当前最新序列号，丢弃迟到响应
      if (currentSeq < refreshSeq) {
        return;
      }

      if (res.ok && res.data) {
        // 刷新成功：更新数据、样本与采集时间，保留用户现有的勾选项
        selectionContext = res.data;
        refreshError = '';
      } else {
        // 刷新失败：保留旧快照数据与旧采集时间，仅展示失败提示
        refreshError = res.error || '刷新原区域失败';
      }
    } catch (err: any) {
      if (selectionContext && selectionContext.attachmentId === currentAttId && currentSeq >= refreshSeq) {
        refreshError = err.message || '刷新原区域异常';
      }
    } finally {
      if (currentSeq >= refreshSeq) {
        isRefreshingArea = false;
      }
    }
  }

  function removeSelection() {
    selectionContext = null;
    selectionError = '';
    refreshError = '';
    showDetails = false;
    ++refreshSeq; // 废弃正在途中的刷新响应
  }

  async function handleSubmit(overrideMode?: 'AUTOMATION' | 'CHAT') {
    if (!inputText.trim() || disabled || isRefreshingArea) return;
    const text = inputText.trim();
    const modeToSend = overrideMode || currentMode;
    if (overrideMode && overrideMode !== currentMode) {
      currentMode = overrideMode;
    }

    try {
      let snapshot: SelectionSnapshot | undefined = undefined;
      if (selectionContext) {
        const { promptText, auditSummary } = formatSelectionContextForPrompt(selectionContext, sendOptions);
        snapshot = {
          context: selectionContext,
          options: { ...sendOptions },
          formattedText: promptText,
          summary: auditSummary,
        };
      }

      const macroRefToSend = referencedMacro;
      const sent = await onSend(text, modeToSend, snapshot, macroRefToSend || undefined);
      if (sent === false) {
        // 被前置门禁阻止：保留输入文字、选区附件与已选引用，不无提示清空！
        return;
      }

      // 只有本次发送被可靠捕获并启动后，才清空输入框并单次解除选区与引用宏，绝不污染后续轮次
      inputText = '';
      selectionContext = null;
      referencedMacro = null;
      showMacroCodeModal = false;
      showDetails = false;
      refreshError = '';
      ++refreshSeq;
    } catch (err: any) {
      console.error('[ChatInput] 组装发送请求异常:', err);
      selectionError = '发送请求异常: ' + (err?.message || err);
    }
  }

  let textareaEl: HTMLTextAreaElement | null = null;

  function handleKeyDown(e: KeyboardEvent) {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSubmit();
    }
  }

  function setMode(mode: 'AUTOMATION' | 'CHAT') {
    currentMode = mode;
  }

  export function setDraftText(val: string, mode: 'AUTOMATION' | 'CHAT' = 'CHAT') {
    inputText = val;
    currentMode = mode;
    setTimeout(() => {
      if (textareaEl) {
        textareaEl.focus();
        textareaEl.setSelectionRange(textareaEl.value.length, textareaEl.value.length);
      }
    }, 50);
  }
</script>

<svelte:window
  on:click={(e) => {
    if (showPlusMenu && plusContainerEl && e.target instanceof Node && !plusContainerEl.contains(e.target)) {
      showPlusMenu = false;
    }
  }}
/>

<div class="input-container">

  <!-- 未存盘工作簿前置预警条 -->
  {#if isUnsavedWorkbook}
    <div class="unsaved-warning-bar" title="当前活动工作簿尚未保存到磁盘">
      <div class="unsaved-msg-wrapper">
        <span class="unsaved-dot">💾</span>
        <span class="unsaved-text">工作簿【{workbook?.name || '未保存'}】未存盘，建议自动化操作前先按 Ctrl+S 存盘。</span>
      </div>
    </div>
  {/if}

  <!-- 选区读取错误提示条 -->
  {#if selectionError}
    <div class="selection-error-bar">
      <div class="error-msg-wrapper">
        <span class="error-dot">⚠️</span>
        <span class="error-text">{selectionError}</span>
      </div>
      <button class="icon-close-btn" on:click={() => (selectionError = '')} title="关闭提示">✕</button>
    </div>
  {/if}

  <!-- 显式引用的前序宏卡片 (用户显式主动选择，单次有效) -->
  {#if referencedMacro}
    <div class="macro-ref-card">
      <div class="card-header">
        <div class="card-title-group">
          <span class="badge-tag macro-badge">🏷️ 已引用前序宏</span>
          <span class="macro-title" title="{referencedMacro.procedureName || '未命名宏'}">
            <strong>{referencedMacro.procedureName || 'Sub Main'}</strong>
          </span>
          <span class="dim-badge">
            {referencedMacro.charCount} 字符 ({referencedMacro.lineCount} 行)
          </span>
          {#if referencedMacro.status === 'incomplete'}
            <span class="status-pill status-warn">代码结构未闭合</span>
          {/if}
          <span class="single-use-badge">仅对本次发送有效</span>
        </div>
        <div class="card-actions-group">
          <button class="text-action-btn" type="button" on:click={() => (showMacroCodeModal = !showMacroCodeModal)}>
            {showMacroCodeModal ? '收起源码 ▲' : '查看源码 ▼'}
          </button>
          <button class="text-action-btn remove-btn" type="button" on:click={() => (referencedMacro = null)} title="取消引用此宏">
            ✕
          </button>
        </div>
      </div>
      {#if showMacroCodeModal}
        <div class="macro-code-box">
          <pre><code>{referencedMacro.vbaCode}</code></pre>
        </div>
      {/if}
    </div>
  {/if}

  <!-- 紧凑选区预览卡片 (点击状态条展开时展示) -->
  {#if selectionContext && showDetails}
    <div class="selection-card">
      <div class="card-header">
        <div class="card-title-group">
          <span class="badge-tag">📎 选区快照</span>
          <span class="target-wb-sheet" title="{selectionContext.workbookName} - {selectionContext.sheetName}!{selectionContext.address}">
            <strong>{selectionContext.sheetName}</strong>!{selectionContext.address}
          </span>
          <span class="dim-badge">
            {selectionContext.totalRows}行 × {selectionContext.totalColumns}列
          </span>
          {#if selectionContext.capturedAt}
            <span class="time-badge" title="采集时系统时间: {formatCaptureTime(selectionContext.capturedAt)}">
              {formatCaptureTime(selectionContext.capturedAt)} 采集
            </span>
          {/if}
        </div>
        <div class="card-actions-group">
          {#if selectionContext.formulaStatus === 'has_formula' || selectionContext.address}
            <button
              class="text-action-btn audit-btn {isAuditingFormula ? 'refreshing' : ''}"
              type="button"
              on:click={handleAuditCurrentFormula}
              disabled={isAuditingFormula || disabled}
              title="向上追溯单元格公式输入源，向下分析从属影响与计算错误"
            >
              <span>{isAuditingFormula ? '体检中...' : '🔍 公式血缘'}</span>
            </button>
          {/if}
          <button class="text-action-btn" type="button" on:click={() => (showDetails = !showDetails)}>
            {showDetails ? '收起详情 ▲' : '查看/设置 ▼'}
          </button>
          <!-- 刷新原区域按钮：定向读取该附件保存的原工作簿/表/地址，不影响当前活动选区 -->
          <button
            class="text-action-btn refresh-btn {isRefreshingArea ? 'refreshing' : ''}"
            type="button"
            on:click={handleRefreshOriginalArea}
            disabled={isRefreshingArea || disabled}
            title="读取并刷新原区域（{selectionContext.sheetName}!{selectionContext.address}）最新数据，不改变当前光标"
          >
            <span class="refresh-icon {isRefreshingArea ? 'spin' : ''}">🔄</span>
            <span>{isRefreshingArea ? '读取中...' : '刷新原区域'}</span>
          </button>
          <button class="text-action-btn remove-btn" type="button" on:click={removeSelection} disabled={isRefreshingArea} title="移除本次附加">
            ✕
          </button>
        </div>
      </div>

      <!-- 快照提示条与错误状态 -->
      <div class="snapshot-status-bar">
        <span class="snapshot-tip-text">ℹ️ 采集时快照；修改原区域后请点击【刷新原区域】</span>
        {#if refreshError}
          <span class="refresh-err-text" title={refreshError}>⚠️ {refreshError} (旧数据仍保留)</span>
        {/if}
      </div>

      <!-- 发送选项复选框 -->
      <div class="card-options">
        <label class="option-label" title="首行是否作为候选表头">
          <input type="checkbox" bind:checked={sendOptions.firstRowAsHeader} />
          <span>首行为表头</span>
        </label>
        <label class="option-label" title="是否发送前几行样本数据值（默认仅发送结构与坐标）">
          <input type="checkbox" bind:checked={sendOptions.includeSamples} />
          <span class={sendOptions.includeSamples ? 'highlight-text' : ''}>含样本值 (前{selectionContext.sampleRowCount}行)</span>
        </label>
        <label class="option-label {sendOptions.includeSamples ? '' : 'disabled-opt'}" title="是否在样本中附加单元格公式（如 =SUM(...)）">
          <input type="checkbox" bind:checked={sendOptions.includeFormulas} disabled={!sendOptions.includeSamples} />
          <span>含样本公式</span>
        </label>
      </div>

      <!-- 展开的详细信息与样本预览 -->
      {#if showDetails}
        <div class="card-details-panel">
          <div class="meta-row">
            <span class="meta-item">工作簿: <code>{selectionContext.workbookName}</code></span>
            <span class="meta-item">公式状态: <code>{selectionContext.formulaStatus}</code></span>
            <span class="meta-item">合并单元格: <code>{selectionContext.mergeStatus}</code></span>
          </div>
          <div class="unscanned-tip">
            ℹ️ {selectionContext.unscannedNotes}
          </div>

          {#if sendOptions.includeSamples && selectionContext.sampleRows && selectionContext.sampleRows.length > 0}
            <div class="sample-table-container">
              <table class="sample-table">
                <thead>
                  <tr>
                    <th>坐标</th>
                    {#each selectionContext.candidateHeaders as h}
                      <th>{h}</th>
                    {/each}
                  </tr>
                </thead>
                <tbody>
                  {#each selectionContext.sampleRows as row}
                    <tr>
                      <td class="coord-cell">{row[0]?.address || ''}</td>
                      {#each row as cell}
                        <td title="值: {cell.displayText}{cell.formula ? ' [公式: ' + cell.formula + ']' : ''}">
                          {cell.displayText || '(空)'}
                          {#if sendOptions.includeFormulas && cell.formula}
                            <span class="cell-formula-tag">fx</span>
                          {/if}
                        </td>
                      {/each}
                    </tr>
                  {/each}
                </tbody>
              </table>
            </div>
          {:else}
            <div class="structure-only-note">
              已选【仅发送结构】：模型仅接收区域地址、行列数与表头结构，不携带单元格具体数值。
            </div>
          {/if}
        </div>
      {/if}
    </div>
  {/if}

  <!-- 输入主框 (压缩单行，单行最小 38px，最大 96px) -->
  <div class="input-box {currentMode === 'AUTOMATION' ? 'focus-auto' : 'focus-chat'}">
    <div class="input-main-row">
      <!-- ＋ 下拉菜单 -->
      <div class="plus-menu-container" bind:this={plusContainerEl}>
        <button
          type="button"
          class="plus-btn"
          on:click={() => (showPlusMenu = !showPlusMenu)}
          disabled={disabled}
          title="添加选区或整簿上下文"
          aria-label="添加选区或整簿上下文"
        >
          <Plus size={14} />
        </button>
        {#if showPlusMenu}
          <div class="plus-dropdown" role="menu">
            <button
              type="button"
              class="plus-dropdown-item"
              on:click={() => { showPlusMenu = false; handleAttachSelection(); }}
              disabled={disabled || isLoadingSelection}
            >
              附加当前选区
            </button>
            <button
              type="button"
              class="plus-dropdown-item"
              on:click={() => { showPlusMenu = false; handleScanWorkbookOverview(); }}
              disabled={disabled || isScanningOverview}
            >
              整簿上下文
            </button>
          </div>
        {/if}
      </div>

      <!-- 输入框 -->
      <textarea
        bind:this={textareaEl}
        placeholder={currentMode === 'AUTOMATION'
          ? "输入操作指令 (如: 用_分裂D列的名称)..."
          : "向 AI 咨询 Excel 问题或技巧..."}
        bind:value={inputText}
        on:keydown={handleKeyDown}
        disabled={disabled}
        rows="1"
      ></textarea>

      <!-- 执行按钮 -->
      <button
        class="send-primary-btn {currentMode === 'AUTOMATION' ? 'send-auto' : 'send-chat'}"
        on:click={() => handleSubmit()}
        disabled={disabled || !inputText.trim() || isRefreshingArea}
        title={isRefreshingArea
          ? "选区正在刷新原区域，请稍候完成后再发送"
          : (currentMode === 'AUTOMATION' ? "以【操作】模式发送：生成自动化方案并修改表格" : "以【对话】模式发送：仅解答咨询，不改动表格")}
        type="button"
      >
        {#if currentMode === 'AUTOMATION'}
          <Zap size={13} />
          <span>执行</span>
        {:else}
          <MessageSquare size={13} />
          <span>发送</span>
        {/if}
      </button>
    </div>
  </div>

  <!-- 状态条 (12px) -->
  <div class="chat-status-bar">
    <div class="status-left">
      <button
        type="button"
        class="status-mode-btn"
        on:click={() => setMode(currentMode === 'AUTOMATION' ? 'CHAT' : 'AUTOMATION')}
        disabled={disabled}
        title="点击切换操作/对话模式"
      >
        {currentMode === 'AUTOMATION' ? '操作' : '对话'}
      </button>
      <span class="status-sep">·</span>
      {#if !selectionContext}
        <button
          type="button"
          class="status-attach-btn"
          on:click={handleAttachSelection}
          disabled={disabled || isLoadingSelection || isRefreshingArea}
          title="点击附加当前 Excel 选区"
        >
          {isLoadingSelection ? '正在读取...' : '未附加选区'}
        </button>
      {:else}
        <button
          type="button"
          class="status-attach-btn attached"
          on:click={() => (showDetails = !showDetails)}
          title="已附加选区，点击展开详情"
        >
          已附加 {selectionContext.sheetName ? `${selectionContext.sheetName}!` : ''}{selectionContext.address}
        </button>
        <button
          type="button"
          class="status-remove-btn"
          on:click={removeSelection}
          disabled={isRefreshingArea}
          title="移除选区"
        >
          ✕
        </button>
      {/if}
    </div>
    <div class="status-right">
      <span class="status-shortcut-info" title="Enter 发送 / Shift+Enter 换行">
        <Info size={12} />
      </span>
    </div>
  </div>

  <!-- 公式血缘穿透审计弹窗 -->
  {#if showAuditModal && auditModalData}
    <div
      class="formula-audit-overlay"
      role="button"
      tabindex="-1"
      on:click|self={() => (showAuditModal = false)}
      on:keydown={(e) => { if (e.key === 'Escape') showAuditModal = false; }}
    >
      <div class="formula-audit-dialog" role="dialog" aria-modal="true" aria-labelledby="audit-dialog-title">
        <div class="dialog-header">
          <div class="dialog-title-group">
            <span class="audit-icon">🔍</span>
            <h3 id="audit-dialog-title">公式血缘穿透审计</h3>
            <span class="cell-tag">{auditModalData.sheetName}!{auditModalData.cellAddress}</span>
          </div>
          <button class="icon-close-btn" type="button" on:click={() => (showAuditModal = false)} title="关闭">✕</button>
        </div>

        <div class="dialog-body">
          <!-- 公式原貌与计算状态 -->
          <div class="audit-formula-box">
            <div class="formula-line">
              <span class="label">公式：</span>
              <code class="formula-text">{auditModalData.formula}</code>
            </div>
            <div class="eval-line">
              <span class="label">计算值：</span>
              <span class="val-text">{auditModalData.formattedText || auditModalData.evaluatedValue}</span>
              {#if auditModalData.hasError}
                <span class="badge-err">⚠️ 错误: {auditModalData.errorType}</span>
              {:else if auditModalData.isCircular}
                <span class="badge-warn">🔄 循环引用风险</span>
              {:else}
                <span class="badge-ok">✓ 结果正常</span>
              {/if}
            </div>
          </div>

          <!-- 上游依赖列表 -->
          <div class="audit-section">
            <div class="section-title">
              <span>上游依赖输入源 ({auditModalData.directPrecedents.length} 处)：</span>
            </div>
            {#if auditModalData.directPrecedents.length === 0}
              <div class="empty-tip">未发现直接上游单元格（常数或独立运算）</div>
            {:else}
              <div class="nodes-list">
                {#each auditModalData.directPrecedents as node}
                  <div class="node-item {node.hasError ? 'node-err' : ''}">
                    <span class="node-addr">{node.sheetName ? `${node.sheetName}!` : ''}{node.address}</span>
                    <span class="node-val" title="当前取值: {node.value}">值: {node.value || '(空)'}</span>
                    {#if node.hasError}
                      <span class="node-badge-err">异常: {node.errorType}</span>
                    {/if}
                  </div>
                {/each}
              </div>
            {/if}
          </div>

          <!-- 下游从属引用 -->
          <div class="audit-section">
            <div class="section-title">
              <span>下游从属引用 ({auditModalData.directDependents.length} 处)：</span>
            </div>
            {#if auditModalData.directDependents.length === 0}
              <div class="empty-tip">暂无下游单元格直接引用此结果</div>
            {:else}
              <div class="dependents-chips">
                {#each auditModalData.directDependents as dep}
                  <span class="dep-chip">{dep}</span>
                {/each}
              </div>
            {/if}
          </div>

          <div class="audit-summary-note">
            ℹ️ {auditModalData.auditSummary}
          </div>
        </div>

        <div class="dialog-footer">
          <button class="action-btn-primary" type="button" on:click={handleInsertExplanationPrompt}>
            <span>🤖 让 AI 白话解读此公式</span>
          </button>
          <button class="action-btn-secondary" type="button" on:click={() => (showAuditModal = false)}>
            关闭
          </button>
        </div>
      </div>
    </div>
  {/if}
</div>

<style>
  .input-container {
    background: #ffffff;
    border-top: 1px solid var(--office-border);
    padding: 4px 8px 6px;
    display: flex;
    flex-direction: column;
    gap: 4px;
    box-shadow: 0 -1px 3px rgba(0, 0, 0, 0.03);
    flex-shrink: 0;
  }

  @keyframes rotate {
    from { transform: rotate(0deg); }
    to { transform: rotate(360deg); }
  }

  /* 未存盘预警条 */
  .unsaved-warning-bar {
    background: #fff8f0;
    border: 1px solid #fed9cc;
    color: #a14c00;
    padding: 3px 8px;
    border-radius: var(--office-radius-sm);
    font-size: 11px;
    display: flex;
    align-items: center;
    gap: 6px;
  }

  .unsaved-msg-wrapper {
    display: flex;
    align-items: center;
    gap: 4px;
    overflow: hidden;
  }

  .unsaved-dot {
    font-size: 11px;
    flex-shrink: 0;
  }

  .unsaved-text {
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    line-height: 1.3;
  }

  /* 选区错误提示条 */
  .selection-error-bar {
    background: #fde7e9;
    border: 1px solid #f9c2c6;
    color: #a80000;
    padding: 4px 8px;
    border-radius: var(--office-radius-sm);
    font-size: 11px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 6px;
  }

  .error-msg-wrapper {
    display: flex;
    align-items: center;
    gap: 4px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .icon-close-btn {
    border: none;
    background: transparent;
    color: #a80000;
    cursor: pointer;
    font-size: 11px;
    padding: 0 2px;
  }

  /* 紧凑只读选区卡片 */
  .selection-card {
    background: #f8faf9;
    border: 1px solid #c7e0d2;
    border-radius: var(--office-radius);
    padding: 6px 8px;
    display: flex;
    flex-direction: column;
    gap: 4px;
    font-size: 11px;
  }

  .card-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 6px;
  }

  .card-title-group {
    display: flex;
    align-items: center;
    gap: 5px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .badge-tag {
    background: #107c41;
    color: white;
    padding: 1px 4px;
    border-radius: 3px;
    font-size: 10px;
    font-weight: 600;
  }

  .target-wb-sheet {
    color: #201f1e;
    font-family: var(--font-family-mono, monospace);
    font-size: 11px;
  }

  .dim-badge {
    background: #edebe9;
    color: #605e5c;
    padding: 1px 4px;
    border-radius: 3px;
    font-size: 10px;
  }

  .time-badge {
    background: #e7f3ec;
    color: #107c41;
    padding: 1px 4px;
    border-radius: 3px;
    font-size: 10px;
    font-family: var(--font-family-mono, monospace);
  }

  .card-actions-group {
    display: flex;
    align-items: center;
    gap: 4px;
    flex-shrink: 0;
  }

  .text-action-btn {
    background: transparent;
    border: none;
    color: #0078d4;
    font-size: 11px;
    cursor: pointer;
    padding: 1px 4px;
    border-radius: 2px;
  }

  .text-action-btn:hover:not(:disabled) {
    background: #e1dfdd;
  }

  .refresh-btn {
    display: inline-flex;
    align-items: center;
    gap: 3px;
    background: #eef6f1;
    color: #107c41;
    font-weight: 500;
    padding: 1px 5px;
    border-radius: 3px;
    border: 1px solid #c7e0d2;
  }

  .refresh-btn:hover:not(:disabled) {
    background: #dff0e6;
  }

  .audit-btn {
    display: inline-flex;
    align-items: center;
    gap: 3px;
    background: #f0fdf4;
    color: #15803d;
    font-weight: 500;
    padding: 1px 5px;
    border-radius: 3px;
    border: 1px solid #bbf7d0;
  }

  .audit-btn:hover:not(:disabled) {
    background: #dcfce7;
    border-color: #86efac;
  }

  .audit-btn:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }

  .refresh-btn:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }

  .refresh-icon {
    display: inline-block;
    font-size: 10px;
  }

  .snapshot-status-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 6px;
    font-size: 10px;
    color: #605e5c;
    background: #f3f5f4;
    padding: 2px 6px;
    border-radius: 2px;
  }

  .snapshot-tip-text {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .refresh-err-text {
    color: #a80000;
    font-weight: 600;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .remove-btn {
    color: #a80000;
    font-weight: bold;
  }

  .card-options {
    display: flex;
    align-items: center;
    gap: 12px;
    padding-top: 3px;
    border-top: 1px dashed #d5e5dc;
    flex-wrap: wrap;
  }

  .option-label {
    display: inline-flex;
    align-items: center;
    gap: 3px;
    cursor: pointer;
    color: #323130;
    user-select: none;
  }

  .highlight-text {
    color: #107c41;
    font-weight: 600;
  }

  .disabled-opt {
    opacity: 0.5;
    cursor: not-allowed;
  }

  /* 展开详情区 */
  .card-details-panel {
    margin-top: 4px;
    padding-top: 4px;
    border-top: 1px solid #d5e5dc;
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  .meta-row {
    display: flex;
    align-items: center;
    gap: 10px;
    flex-wrap: wrap;
    color: #605e5c;
  }

  .meta-item code {
    background: #f0f0f0;
    padding: 1px 3px;
    border-radius: 2px;
    font-size: 10px;
  }

  .unscanned-tip {
    color: #797775;
    font-size: 10px;
    line-height: 1.3;
  }

  .structure-only-note {
    background: #f3f2f1;
    color: #605e5c;
    padding: 4px 6px;
    border-radius: 3px;
    font-style: italic;
  }

  .sample-table-container {
    max-height: 120px;
    overflow: auto;
    border: 1px solid #edebe9;
    border-radius: 3px;
    background: white;
  }

  .sample-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 10px;
    text-align: left;
  }

  .sample-table th, .sample-table td {
    padding: 3px 6px;
    border: 1px solid #edebe9;
    white-space: nowrap;
    max-width: 120px;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  .sample-table th {
    background: #f8f8f8;
    color: #323130;
    font-weight: 600;
    position: sticky;
    top: 0;
  }

  .coord-cell {
    background: #fdfdfd;
    color: #605e5c;
    font-weight: 500;
  }

  .cell-formula-tag {
    display: inline-block;
    background: #0078d4;
    color: white;
    font-size: 8px;
    padding: 0 2px;
    border-radius: 2px;
    margin-left: 2px;
  }

  /* 输入主框 (压缩单行，单行最小 38px，最大 96px) */
  .input-box {
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    background: #ffffff;
    display: flex;
    flex-direction: column;
    transition: all 0.15s ease;
    padding: 2px 6px;
    box-sizing: border-box;
  }

  .input-box.focus-auto:focus-within {
    border-color: var(--excel-green);
    box-shadow: 0 0 0 2px rgba(16, 124, 65, 0.15);
  }

  .input-box.focus-chat:focus-within {
    border-color: var(--office-blue);
    box-shadow: 0 0 0 2px rgba(0, 120, 212, 0.15);
  }

  .input-main-row {
    display: flex;
    align-items: center;
    gap: 6px;
    width: 100%;
    box-sizing: border-box;
  }

  /* ＋ 菜单 */
  .plus-menu-container {
    position: relative;
    display: inline-flex;
    align-items: center;
    flex-shrink: 0;
  }

  .plus-btn {
    width: 26px;
    height: 26px;
    border-radius: var(--office-radius-xs);
    border: 1px solid var(--office-border-subtle);
    background: #f8f9fa;
    color: var(--office-text-secondary);
    display: flex;
    align-items: center;
    justify-content: center;
    cursor: pointer;
    transition: all 0.15s ease;
    padding: 0;
  }

  .plus-btn:hover:not(:disabled) {
    background: var(--office-hover);
    color: var(--office-text);
    border-color: var(--office-border);
  }

  .plus-btn:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }

  .plus-dropdown {
    position: absolute;
    bottom: 100%;
    left: 0;
    margin-bottom: 6px;
    background: #ffffff;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-sm);
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
    padding: 4px;
    display: flex;
    flex-direction: column;
    gap: 2px;
    z-index: 50;
    min-width: 110px;
  }

  .plus-dropdown-item {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 5px 8px;
    font-size: var(--font-size-xs);
    color: var(--office-text);
    border: none;
    background: transparent;
    border-radius: var(--office-radius-xs);
    cursor: pointer;
    text-align: left;
    white-space: nowrap;
    transition: background 0.12s ease;
  }

  .plus-dropdown-item:hover:not(:disabled) {
    background: var(--office-hover);
    color: var(--excel-green);
  }

  .plus-dropdown-item:disabled {
    opacity: 0.5;
    cursor: not-allowed;
  }

  textarea {
    flex: 1 1 0%;
    min-width: 0;
    width: 100%;
    resize: none;
    border: none;
    outline: none;
    padding: 8px 4px;
    font-size: var(--font-size-base, 13px);
    font-family: var(--font-family-ui);
    line-height: var(--line-height-normal);
    color: var(--office-text);
    min-height: 38px;
    max-height: 96px;
    background: transparent;
    overflow-y: auto;
    box-sizing: border-box;
  }

  textarea:disabled {
    background: #faf9f8;
    color: var(--office-muted);
  }

  .send-primary-btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 4px;
    padding: 0 10px;
    height: 30px;
    font-size: var(--font-size-xs, 12px);
    font-weight: 600;
    border-radius: var(--office-radius-sm);
    cursor: pointer;
    transition: all 0.15s ease;
    white-space: nowrap;
    border: 1px solid transparent;
    flex-shrink: 0;
  }

  .send-auto {
    background: var(--excel-green);
    border-color: var(--excel-green);
    color: #ffffff;
  }

  .send-auto:hover:not(:disabled) {
    background: var(--excel-hover-bg);
    border-color: var(--excel-hover-bg);
  }

  .send-chat {
    background: var(--office-blue);
    border-color: var(--office-blue);
    color: #ffffff;
  }

  .send-chat:hover:not(:disabled) {
    background: var(--office-blue-dark);
    border-color: var(--office-blue-dark);
  }

  .send-primary-btn:disabled {
    border-color: var(--office-dim);
    opacity: 0.6;
    cursor: not-allowed;
  }

  /* 状态条 (12px) */
  .chat-status-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    height: 18px;
    font-size: 12px;
    color: var(--office-dim);
    padding: 0 2px;
    margin-top: 2px;
    box-sizing: border-box;
  }

  .status-left {
    display: flex;
    align-items: center;
    gap: 4px;
    min-width: 0;
    overflow: hidden;
  }

  .status-mode-btn {
    background: transparent;
    border: none;
    font-size: 12px;
    font-weight: 600;
    color: var(--office-text-secondary);
    cursor: pointer;
    padding: 0 2px;
    line-height: 1;
    transition: color 0.15s ease;
  }

  .status-mode-btn:hover:not(:disabled) {
    color: var(--excel-green);
  }

  .status-sep {
    color: var(--office-dim);
    font-size: 11px;
    user-select: none;
  }

  .status-attach-btn {
    background: transparent;
    border: none;
    font-size: 12px;
    color: var(--office-dim);
    cursor: pointer;
    padding: 0 2px;
    line-height: 1;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 220px;
    transition: color 0.15s ease;
  }

  .status-attach-btn:hover:not(:disabled) {
    color: var(--office-text);
  }

  .status-attach-btn.attached {
    color: var(--excel-green);
    font-weight: 500;
  }

  .status-remove-btn {
    background: transparent;
    border: none;
    color: var(--office-dim);
    font-size: 10px;
    cursor: pointer;
    padding: 0 2px;
    line-height: 1;
  }

  .status-remove-btn:hover:not(:disabled) {
    color: var(--office-red, #d13438);
  }

  .status-right {
    display: flex;
    align-items: center;
    flex-shrink: 0;
  }

  .status-shortcut-info {
    display: flex;
    align-items: center;
    color: var(--office-dim);
    cursor: help;
  }

  .status-shortcut-info:hover {
    color: var(--office-text);
  }

  /* 操作模式发送按钮：Excel 绿 */
  .send-auto {
    background: var(--excel-green);
    border-color: var(--excel-green);
    color: #ffffff;
    box-shadow: 0 1px 2px rgba(16, 124, 65, 0.2);
  }

  .send-auto:hover:not(:disabled) {
    background: var(--excel-hover-bg);
    border-color: var(--excel-hover-bg);
  }

  /* 对话模式发送按钮：Office 蓝 */
  .send-chat {
    background: var(--office-blue);
    border-color: var(--office-blue);
    color: #ffffff;
    box-shadow: 0 1px 2px rgba(0, 120, 212, 0.2);
  }

  .send-chat:hover:not(:disabled) {
    background: var(--office-blue-dark);
    border-color: var(--office-blue-dark);
  }

  .macro-ref-card {
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: var(--office-radius, 6px);
    padding: 8px 12px;
    margin-bottom: 8px;
    display: flex;
    flex-direction: column;
    gap: 6px;
    font-size: var(--font-size-xs, 12px);
  }
  .macro-badge {
    background: #e0f2fe;
    color: #0369a1;
    font-size: 11px;
    padding: 2px 6px;
    border-radius: 4px;
    font-weight: 500;
  }
  .macro-title {
    font-size: 12px;
    color: #1e293b;
  }
  .single-use-badge {
    font-size: 11px;
    color: #64748b;
  }
  .status-pill.status-warn {
    font-size: 10px;
    background: #fef3c7;
    color: #b45309;
    padding: 1px 6px;
    border-radius: 4px;
  }
  .macro-code-box {
    background: #1e1e1e;
    color: #d4d4d4;
    padding: 8px 10px;
    border-radius: 4px;
    font-size: 11px;
    max-height: 180px;
    overflow-y: auto;
  }
  .macro-code-box pre {
    margin: 0;
    white-space: pre-wrap;
    word-break: break-all;
    font-family: Consolas, monospace;
  }

  /* 公式血缘穿透审计弹窗样式 */
  .formula-audit-overlay {
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background: rgba(0, 0, 0, 0.45);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 1000;
    backdrop-filter: blur(2px);
  }

  .formula-audit-dialog {
    background: #ffffff;
    border-radius: 8px;
    width: 90%;
    max-width: 480px;
    box-shadow: 0 10px 25px rgba(0, 0, 0, 0.2);
    display: flex;
    flex-direction: column;
    max-height: 85vh;
    animation: fadeIn 0.2s ease-out;
  }

  @keyframes fadeIn {
    from { opacity: 0; transform: scale(0.96); }
    to { opacity: 1; transform: scale(1); }
  }

  .dialog-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 10px 14px;
    border-bottom: 1px solid #e2e8f0;
  }

  .dialog-title-group {
    display: flex;
    align-items: center;
    gap: 6px;
  }

  .dialog-title-group h3 {
    margin: 0;
    font-size: 13px;
    font-weight: 600;
    color: #0f172a;
  }

  .cell-tag {
    background: #eff6ff;
    color: #1d4ed8;
    font-size: 11px;
    font-family: monospace;
    padding: 1px 6px;
    border-radius: 4px;
    border: 1px solid #bfdbfe;
  }

  .dialog-body {
    padding: 12px 14px;
    overflow-y: auto;
    display: flex;
    flex-direction: column;
    gap: 12px;
    font-size: 12px;
  }

  .audit-formula-box {
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 8px 10px;
    display: flex;
    flex-direction: column;
    gap: 6px;
  }

  .formula-line {
    display: flex;
    align-items: baseline;
    gap: 4px;
  }

  .formula-line .label, .eval-line .label {
    color: #64748b;
    font-size: 11px;
    flex-shrink: 0;
  }

  .formula-text {
    color: #0972d3;
    font-family: Consolas, monospace;
    font-size: 12px;
    font-weight: bold;
    word-break: break-all;
  }

  .eval-line {
    display: flex;
    align-items: center;
    gap: 6px;
  }

  .val-text {
    font-weight: 600;
    color: #0f172a;
  }

  .badge-err {
    background: #fef2f2;
    color: #b91c1c;
    font-size: 10px;
    padding: 1px 6px;
    border-radius: 4px;
    border: 1px solid #fecaca;
  }

  .badge-warn {
    background: #fffbeb;
    color: #b45309;
    font-size: 10px;
    padding: 1px 6px;
    border-radius: 4px;
    border: 1px solid #fde68a;
  }

  .badge-ok {
    background: #f0fdf4;
    color: #15803d;
    font-size: 10px;
    padding: 1px 6px;
    border-radius: 4px;
    border: 1px solid #bbf7d0;
  }

  .audit-section {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }

  .section-title {
    font-size: 11px;
    font-weight: 600;
    color: #475569;
  }

  .nodes-list {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  .node-item {
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #f1f5f9;
    padding: 4px 8px;
    border-radius: 4px;
    font-size: 11px;
  }

  .node-item.node-err {
    background: #fef2f2;
    border: 1px solid #fecaca;
  }

  .node-addr {
    font-family: monospace;
    font-weight: 600;
    color: #334155;
  }

  .node-val {
    color: #64748b;
  }

  .node-badge-err {
    color: #dc2626;
    font-weight: 500;
  }

  .dependents-chips {
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
  }

  .dep-chip {
    background: #e2e8f0;
    color: #334155;
    font-family: monospace;
    font-size: 11px;
    padding: 2px 6px;
    border-radius: 4px;
  }

  .empty-tip {
    color: #94a3b8;
    font-size: 11px;
    font-style: italic;
  }

  .audit-summary-note {
    background: #eff6ff;
    border: 1px solid #dbeafe;
    border-radius: 4px;
    padding: 6px 8px;
    font-size: 11px;
    color: #1e40af;
    line-height: 1.4;
  }

  .dialog-footer {
    display: flex;
    align-items: center;
    justify-content: flex-end;
    gap: 8px;
    padding: 10px 14px;
    border-top: 1px solid #e2e8f0;
    background: #fafafa;
    border-radius: 0 0 8px 8px;
  }

  .action-btn-primary {
    background: #0078d4;
    color: white;
    border: none;
    padding: 5px 12px;
    border-radius: 4px;
    font-size: 12px;
    font-weight: 500;
    cursor: pointer;
    transition: background 0.15s;
  }

  .action-btn-primary:hover {
    background: #106ebe;
  }

  .action-btn-secondary {
    background: #ffffff;
    border: 1px solid #cbd5e1;
    color: #334155;
    padding: 5px 10px;
    border-radius: 4px;
    font-size: 12px;
    cursor: pointer;
  }

  .action-btn-secondary:hover {
    background: #f1f5f9;
  }
</style>
