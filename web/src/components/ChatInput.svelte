<script lang="ts">
  import { Zap, MessageSquare } from 'lucide-svelte';
  import { bridge, type SelectionContextData, type SelectionSendOptions, type SelectionSnapshot } from '../services/bridge';
  import { formatSelectionContextForPrompt } from '../services/llm';

  export let disabled = false;
  export let onSend: (text: string, mode?: 'AUTOMATION' | 'CHAT', selectionSnapshot?: SelectionSnapshot) => void;

  let inputText = '';
  let currentMode: 'AUTOMATION' | 'CHAT' = 'AUTOMATION';

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

  function handleSubmit(overrideMode?: 'AUTOMATION' | 'CHAT') {
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

      inputText = '';
      onSend(text, modeToSend, snapshot);

      // 发送瞬间冻结并清空本次挂载的选区，保持单次请求固定独立性，后续刷新绝不影响已发送内容
      selectionContext = null;
      showDetails = false;
      refreshError = '';
      ++refreshSeq;
    } catch (err: any) {
      console.error('[ChatInput] 组装发送请求异常:', err);
      selectionError = '发送请求异常: ' + (err?.message || err);
    }
  }

  function handleKeyDown(e: KeyboardEvent) {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSubmit();
    }
  }

  function setMode(mode: 'AUTOMATION' | 'CHAT') {
    currentMode = mode;
  }
</script>

<div class="input-container">
  <!-- 模式切换与附加选区操作栏 -->
  <div class="mode-header">
    <div class="mode-segmented" role="tablist" aria-label="对话模式选择">
      <button
        class="mode-tab {currentMode === 'AUTOMATION' ? 'active-auto' : ''}"
        on:click={() => setMode('AUTOMATION')}
        disabled={disabled}
        type="button"
        role="tab"
        aria-selected={currentMode === 'AUTOMATION'}
        title="操作模式：生成代码并执行，直接修改表格"
      >
        <Zap size={13} class="mode-icon" />
        <span class="mode-title">操作</span>
      </button>
      <button
        class="mode-tab {currentMode === 'CHAT' ? 'active-chat' : ''}"
        on:click={() => setMode('CHAT')}
        disabled={disabled}
        type="button"
        role="tab"
        aria-selected={currentMode === 'CHAT'}
        title="对话模式：纯文本咨询指导，不改动表格"
      >
        <MessageSquare size={13} class="mode-icon" />
        <span class="mode-title">对话</span>
      </button>
    </div>

    <!-- 附加当前选区按钮 (双模式均可用，用于首次添加或替换附件) -->
    <button
      class="attach-selection-btn {selectionContext ? 'has-attached' : ''}"
      on:click={handleAttachSelection}
      disabled={disabled || isLoadingSelection || isRefreshingArea}
      type="button"
      title="只读读取当前 Excel 选中的单元格区域（用于替换现有附件）"
    >
      <svg class="btn-icon {isLoadingSelection ? 'spin' : ''}" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="m21.44 11.05-9.19 9.19a6 6 0 0 1-8.49-8.49l8.57-8.57A4 4 0 1 1 18 8.84l-8.59 8.57a2 2 0 0 1-2.83-2.83l8.49-8.48"/>
      </svg>
      <span>{isLoadingSelection ? '读取中...' : (selectionContext ? '替换为当前选区' : '附加当前选区')}</span>
    </button>
  </div>

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

  <!-- 紧凑选区预览卡片 (当成功附加时展示) -->
  {#if selectionContext}
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

  <!-- 输入主框 (加高输入区域，提供更充足的书写空间) -->
  <div class="input-box {currentMode === 'AUTOMATION' ? 'focus-auto' : 'focus-chat'}">
    <textarea
      placeholder={currentMode === 'AUTOMATION'
        ? "输入操作指令 (如: 用_分裂D列的名称，不要覆盖后面的列，新增)..."
        : "向 AI 咨询 Excel 问题或技巧 (如: TEXTSPLIT与分列区别、公式怎么写)..."}
      bind:value={inputText}
      on:keydown={handleKeyDown}
      disabled={disabled}
      rows="3"
    ></textarea>

    <div class="input-actions">
      <div class="shortcut-tip">
        <span class="key-pill">Enter</span> 发送
        <span class="key-sep">/</span>
        <span class="key-pill">Shift+Enter</span> 换行
      </div>
      <div class="action-buttons-group">
        <!-- 对话按钮 -->
        <button
          class="mode-action-btn {currentMode === 'CHAT' ? 'primary-chat' : 'secondary-chat'}"
          on:click={() => handleSubmit('CHAT')}
          disabled={disabled || !inputText.trim() || isRefreshingArea}
          title={isRefreshingArea ? "选区正在刷新原区域，请稍候完成后再发送" : "以【对话】发送：仅解答，不改动表格"}
          type="button"
        >
          <MessageSquare size={13} />
          <span>对话</span>
        </button>

        <!-- 操作按钮 -->
        <button
          class="mode-action-btn {currentMode === 'AUTOMATION' ? 'primary-auto' : 'secondary-auto'}"
          on:click={() => handleSubmit('AUTOMATION')}
          disabled={disabled || !inputText.trim() || isRefreshingArea}
          title={isRefreshingArea ? "选区正在刷新原区域，请稍候完成后再发送" : "以【操作】发送：直接修改表格并自动备份"}
          type="button"
        >
          <Zap size={13} />
          <span>操作</span>
        </button>
      </div>
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
    gap: 6px;
    box-shadow: 0 -1px 3px rgba(0, 0, 0, 0.03);
    flex-shrink: 0;
  }

  /* 顶部模式切换栏 */
  .mode-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    min-height: 26px;
  }

  .mode-segmented {
    display: inline-flex;
    background: #f0f2f5;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    padding: 2px;
    gap: 2px;
    flex-shrink: 0;
  }

  .mode-tab {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    padding: 3px 10px;
    font-size: var(--font-size-xs);
    font-weight: 500;
    border: none;
    border-radius: var(--office-radius-sm);
    background: transparent;
    color: var(--office-muted);
    cursor: pointer;
    transition: all 0.15s ease;
    line-height: var(--line-height-tight);
  }

  .mode-tab:hover:not(:disabled) {
    color: var(--office-text);
  }

  .mode-tab.active-auto {
    background: var(--excel-green);
    color: #ffffff;
    font-weight: 600;
    box-shadow: 0 1px 3px rgba(16, 124, 65, 0.25);
  }

  .mode-tab.active-chat {
    background: var(--office-blue);
    color: #ffffff;
    font-weight: 600;
    box-shadow: 0 1px 3px rgba(0, 120, 212, 0.25);
  }

  .mode-tab:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }

  /* 附加选区按钮 */
  .attach-selection-btn {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    padding: 3px 8px;
    font-size: var(--font-size-xs);
    font-weight: 500;
    border-radius: var(--office-radius-sm);
    border: 1px solid var(--office-border);
    background: #faf9f8;
    color: #323130;
    cursor: pointer;
    transition: all 0.15s ease;
  }

  .attach-selection-btn:hover:not(:disabled) {
    background: #f3f2f1;
    border-color: #c8c6c4;
  }

  .attach-selection-btn.has-attached {
    background: #e7f3ec;
    border-color: #107c41;
    color: #107c41;
    font-weight: 600;
  }

  .btn-icon {
    flex-shrink: 0;
  }

  .spin {
    animation: rotate 1s linear infinite;
  }

  @keyframes rotate {
    from { transform: rotate(0deg); }
    to { transform: rotate(360deg); }
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

  /* 输入框 (加高设计与舒适边距) */
  .input-box {
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    background: #ffffff;
    display: flex;
    flex-direction: column;
    transition: all 0.15s ease;
  }

  .input-box.focus-auto:focus-within {
    border-color: var(--excel-green);
    box-shadow: 0 0 0 2px rgba(16, 124, 65, 0.15);
  }

  .input-box.focus-chat:focus-within {
    border-color: var(--office-blue);
    box-shadow: 0 0 0 2px rgba(0, 120, 212, 0.15);
  }

  textarea {
    width: 100%;
    resize: none;
    border: none;
    outline: none;
    padding: 8px 10px;
    font-size: var(--font-size-base);
    font-family: var(--font-family-ui);
    line-height: var(--line-height-normal);
    color: var(--office-text);
    min-height: 72px;
    max-height: 180px;
    background: transparent;
    overflow-y: auto;
  }

  textarea:disabled {
    background: #faf9f8;
    color: var(--office-muted);
  }

  /* 底部操作行 */
  .input-actions {
    padding: 5px 8px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #faf9f8;
    border-top: 1px solid var(--office-border-subtle);
    border-bottom-left-radius: calc(var(--office-radius) - 1px);
    border-bottom-right-radius: calc(var(--office-radius) - 1px);
    gap: 8px;
  }

  .shortcut-tip {
    font-size: var(--font-size-xs);
    color: var(--office-muted);
    display: flex;
    align-items: center;
    gap: 3px;
    white-space: nowrap;
    overflow: hidden;
  }

  .key-pill {
    display: inline-block;
    padding: 1px 4px;
    background: #edebe9;
    border-radius: var(--office-radius-xs);
    font-size: 10px;
    font-weight: 500;
    color: #323130;
    line-height: 1.2;
    border: 1px solid #e1dfdd;
  }

  .key-sep {
    color: #c8c6c4;
    margin: 0 1px;
  }

  .action-buttons-group {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    flex-shrink: 0;
  }

  .mode-action-btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 4px;
    padding: 4px 12px;
    font-size: var(--font-size-sm);
    font-weight: 500;
    border-radius: var(--office-radius-sm);
    cursor: pointer;
    transition: all 0.15s ease;
    line-height: 1.3;
    white-space: nowrap;
  }

  .mode-action-btn:disabled {
    opacity: 0.5;
    cursor: not-allowed;
  }

  /* 操作模式主按钮 */
  .primary-auto {
    background: var(--excel-green);
    border: 1px solid var(--excel-green);
    color: #ffffff;
    box-shadow: 0 1px 2px rgba(16, 124, 65, 0.2);
  }

  .primary-auto:hover:not(:disabled) {
    background: var(--excel-hover-bg);
    border-color: var(--excel-hover-bg);
  }

  /* 操作模式次按钮 (当前是对话时) */
  .secondary-auto {
    background: #ffffff;
    border: 1px solid var(--excel-light-border);
    color: var(--excel-green);
  }

  .secondary-auto:hover:not(:disabled) {
    background: var(--excel-light);
  }

  /* 对话模式主按钮 */
  .primary-chat {
    background: var(--office-blue);
    border: 1px solid var(--office-blue);
    color: #ffffff;
    box-shadow: 0 1px 2px rgba(0, 120, 212, 0.2);
  }

  .primary-chat:hover:not(:disabled) {
    background: var(--office-blue-dark);
    border-color: var(--office-blue-dark);
  }

  /* 对话模式次按钮 (当前是操作时) */
  .secondary-chat {
    background: #ffffff;
    border: 1px solid var(--office-blue-border);
    color: var(--office-blue);
  }

  .secondary-chat:hover:not(:disabled) {
    background: var(--office-blue-light);
  }
</style>
