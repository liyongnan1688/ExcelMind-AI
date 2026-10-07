<script lang="ts">
  import { onMount } from 'svelte';
  import {
    X,
    Play,
    Save,
    RotateCcw,
    CheckCircle2,
    AlertCircle,
    Clock,
    Plus,
    Trash2,
    FileSpreadsheet,
    GitBranch,
    StopCircle,
    ShieldAlert,
    ChevronRight,
    ArrowDown
  } from 'lucide-svelte';
  import {
    bridge,
    type WorkbookInfo,
    type WorkflowDefinition,
    type WorkflowStepDefinition,
    type WorkflowRunRecord
  } from '../services/bridge';

  export let isOpen = false;
  export let workbook: WorkbookInfo | null = null;
  export let onClose: () => void;
  export let onExecuted: () => void = () => {};

  let workflows: WorkflowDefinition[] = [];
  let selectedWorkflowId: string = '';
  let activeWorkflow: WorkflowDefinition | null = null;
  let isEditing = false;
  let isExecuting = false;
  let currentRun: WorkflowRunRecord | null = null;
  let runHistory: WorkflowRunRecord[] = [];
  let errorMessage: string = '';
  let successNotice: string = '';

  // 临时编辑模型
  let editModel: WorkflowDefinition = createEmptyWorkflow();

  function createEmptyWorkflow(): WorkflowDefinition {
    return {
      name: '新建双步骤流水线',
      description: '定义两步线性数据处理，全流程任务快照保护',
      definitionVersion: 1,
      targetWorkbookName: workbook?.name || '',
      targetWorkbookPath: workbook?.fullName || '',
      steps: [
        {
          stepIndex: 1,
          stepName: '第 1 步：按键去重',
          toolType: 'dedup',
          inputSource: 'initial_selection',
          dedupParams: {
            targetSheet: workbook?.activeSheetName || 'Sheet1',
            rangeAddress: workbook?.usedRangeAddress || 'A1:D20',
            hasHeader: true,
            keyColumns: [1],
            mode: 'export_unique'
          }
        },
        {
          stepIndex: 2,
          stepName: '第 2 步：主键对账',
          toolType: 'reconcile',
          inputSource: 'prev_step_output',
          reconcileParams: {
            leftHasHeader: true,
            leftKeyCols: [1],
            rightSheet: 'Sheet2',
            rightAddress: 'A1:D20',
            rightHasHeader: true,
            rightKeyCols: [1]
          }
        }
      ]
    };
  }

  onMount(async () => {
    if (isOpen) {
      await loadWorkflows();
    }
  });

  $: if (isOpen) {
    loadWorkflows();
  }

  async function loadWorkflows() {
    errorMessage = '';
    const res = await bridge.listWorkflows();
    if (res.ok && res.data) {
      workflows = res.data;
      if (workflows.length > 0 && !selectedWorkflowId) {
        selectWorkflow(workflows[0].workflowId || '');
      } else if (selectedWorkflowId) {
        selectWorkflow(selectedWorkflowId);
      }
    } else {
      errorMessage = res.error || '获取工作流列表失败';
    }
  }

  async function selectWorkflow(id: string) {
    selectedWorkflowId = id;
    errorMessage = '';
    successNotice = '';
    const res = await bridge.getWorkflow(id);
    if (res.ok && res.data) {
      activeWorkflow = res.data.definition;
      runHistory = res.data.runs || [];
      editModel = JSON.parse(JSON.stringify(activeWorkflow));
      isEditing = false;
      currentRun = null;
    } else {
      errorMessage = res.error || '获取工作流定义失败';
    }
  }

  function startNewWorkflow() {
    activeWorkflow = null;
    selectedWorkflowId = '';
    editModel = createEmptyWorkflow();
    isEditing = true;
    currentRun = null;
    runHistory = [];
  }

  async function handleSave() {
    errorMessage = '';
    successNotice = '';

    if (!editModel.name.trim()) {
      errorMessage = '流水线名称不能为空';
      return;
    }

    if (!editModel.steps || editModel.steps.length !== 2) {
      errorMessage = '必须且仅允许配置 2 个步骤';
      return;
    }

    // 更新绑定的工作簿
    if (workbook?.name && workbook.name !== '未检测到活动工作簿') {
      editModel.targetWorkbookName = workbook.name;
      editModel.targetWorkbookPath = workbook.fullName;
    }

    const res = await bridge.saveWorkflow(editModel);
    if (res.ok && res.data) {
      successNotice = res.message || '工作流保存成功';
      await loadWorkflows();
      if (res.data.workflowId) {
        await selectWorkflow(res.data.workflowId);
      }
      isEditing = false;
    } else {
      errorMessage = res.error || '保存工作流失败';
    }
  }

  async function handleDelete() {
    if (!selectedWorkflowId) return;
    if (!confirm('确定要删除此工作流定义及其运行历史吗？')) return;

    const res = await bridge.deleteWorkflow(selectedWorkflowId);
    if (res.ok) {
      selectedWorkflowId = '';
      activeWorkflow = null;
      await loadWorkflows();
    } else {
      errorMessage = res.error || '删除失败';
    }
  }

  async function handleExecute() {
    if (!activeWorkflow) return;
    if (isExecuting) return;

    isExecuting = true;
    errorMessage = '';
    successNotice = '';
    currentRun = null;

    try {
      const res = await bridge.executeWorkflow({
        workflowId: activeWorkflow.workflowId,
        workflow: activeWorkflow,
        targetWorkbookName: workbook?.name || activeWorkflow.targetWorkbookName,
        targetWorkbookFullName: workbook?.fullName || activeWorkflow.targetWorkbookPath
      });

      if (res.data) {
        currentRun = res.data;
        if (res.ok) {
          successNotice = '流水线执行完成！两步均已成功处理。';
        } else {
          errorMessage = res.data.failureMessage || res.error || '执行未完全成功';
        }
      } else {
        errorMessage = res.error || '未能获取执行响应';
      }

      // 刷新历史
      if (activeWorkflow.workflowId) {
        const histRes = await bridge.getWorkflow(activeWorkflow.workflowId);
        if (histRes.ok && histRes.data) {
          runHistory = histRes.data.runs || [];
        }
      }

      onExecuted();
    } catch (e: any) {
      errorMessage = '执行异常: ' + e.message;
    } finally {
      isExecuting = false;
    }
  }

  async function handleCancel() {
    await bridge.cancelWorkflow();
    errorMessage = '已发送取消指令，等待步骤边界安全生效...';
  }
</script>

{#if isOpen}
  <div class="modal-backdrop" role="presentation" on:click|self={onClose}>
    <div class="modal-window">
      <!-- 头部 -->
      <div class="modal-header">
        <div class="header-title">
          <GitBranch size={16} class="text-indigo-600" />
          <span>双步骤任务流水线 (TASK-R5a)</span>
          {#if activeWorkflow}
            <span class="version-badge">v{activeWorkflow.definitionVersion}</span>
            {#if activeWorkflow.isPreset}
              <span class="preset-tag">内置预设</span>
            {/if}
          {/if}
        </div>
        <button class="close-btn" on:click={onClose} aria-label="关闭">
          <X size={16} />
        </button>
      </div>

      <!-- 主体分栏 -->
      <div class="modal-body-container">
        <!-- 左侧：工作流列表 -->
        <div class="sidebar-list">
          <div class="sidebar-toolbar">
            <span class="toolbar-label">流水线定义库</span>
            <button class="btn-xs add-btn" on:click={startNewWorkflow}>
              <Plus size={12} />
              <span>新建</span>
            </button>
          </div>
          <div class="workflows-scroll">
            {#each workflows as wf}
              <button
                class="wf-item-btn"
                class:active={wf.workflowId === selectedWorkflowId}
                on:click={() => selectWorkflow(wf.workflowId || '')}
              >
                <div class="wf-item-title">
                  <span>{wf.name}</span>
                  <span class="wf-item-ver">v{wf.definitionVersion}</span>
                </div>
                <div class="wf-item-desc">{wf.description || '无描述'}</div>
              </button>
            {/each}
          </div>
        </div>

        <!-- 右侧：详情与执行 -->
        <div class="main-content-panel">
          {#if errorMessage}
            <div class="alert-box error-alert">
              <AlertCircle size={14} />
              <span>{errorMessage}</span>
            </div>
          {/if}

          {#if successNotice}
            <div class="alert-box success-alert">
              <CheckCircle2 size={14} />
              <span>{successNotice}</span>
            </div>
          {/if}

          <!-- 编辑状态 -->
          {#if isEditing}
            <div class="config-form">
              <div class="form-row">
                <span class="form-label">流水线名称</span>
                <input type="text" class="input-text" bind:value={editModel.name} placeholder="例如：销售日报去重并对账" />
              </div>
              <div class="form-row">
                <span class="form-label">业务描述</span>
                <input type="text" class="input-text" bind:value={editModel.description} placeholder="说明业务场景与目的" />
              </div>

              <!-- 第 1 步配置 -->
              <div class="step-config-card">
                <div class="step-card-header">
                  <span class="step-badge">Step 1</span>
                  <input type="text" class="input-step-name" bind:value={editModel.steps[0].stepName} />
                  <select class="select-tool" bind:value={editModel.steps[0].toolType}>
                    <option value="dedup">按键去重工具 (dedup)</option>
                    <option value="saved_macro">已保存宏 (saved_macro)</option>
                  </select>
                </div>
                <div class="step-card-body">
                  {#if editModel.steps[0].toolType === 'dedup' && editModel.steps[0].dedupParams}
                    <div class="param-grid">
                      <div class="param-col">
                        <span class="form-label">目标工作表</span>
                        <input type="text" class="input-text" bind:value={editModel.steps[0].dedupParams.targetSheet} />
                      </div>
                      <div class="param-col">
                        <span class="form-label">单元格区域</span>
                        <input type="text" class="input-text" bind:value={editModel.steps[0].dedupParams.rangeAddress} />
                      </div>
                      <div class="param-col">
                        <span class="form-label">执行模式</span>
                        <select class="input-text" bind:value={editModel.steps[0].dedupParams.mode}>
                          <option value="export_unique">导出唯一行新表</option>
                          <option value="highlight">原表高亮标记</option>
                        </select>
                      </div>
                    </div>
                  {:else if editModel.steps[0].toolType === 'saved_macro' && editModel.steps[0].macroParams}
                    <div class="param-grid">
                      <div class="param-col">
                        <span class="form-label">已保存宏 ID</span>
                        <input type="text" class="input-text" bind:value={editModel.steps[0].macroParams.scriptId} placeholder="如 script_001" />
                      </div>
                      <div class="param-col">
                        <span class="form-label">入口过程</span>
                        <input type="text" class="input-text" bind:value={editModel.steps[0].macroParams.entryPoint} placeholder="如 Sub Main" />
                      </div>
                      <div class="param-col">
                        <span class="form-label">声明输出工作表</span>
                        <input type="text" class="input-text" bind:value={editModel.steps[0].macroParams.declaredOutputSheet} placeholder="宏写出的表名" />
                      </div>
                    </div>
                  {/if}
                </div>
              </div>

              <!-- 管道流动标识 -->
              <div class="pipe-connector">
                <ArrowDown size={14} class="text-indigo-500" />
                <span class="pipe-text">第一步确权输出自动传递为第二步输入源 (prev_step_output)</span>
              </div>

              <!-- 第 2 步配置 -->
              <div class="step-config-card">
                <div class="step-card-header">
                  <span class="step-badge">Step 2</span>
                  <input type="text" class="input-step-name" bind:value={editModel.steps[1].stepName} />
                  <select
                    class="select-tool"
                    bind:value={editModel.steps[1].toolType}
                    on:change={() => {
                      if (editModel.steps[1].toolType === 'chart' && !editModel.steps[1].chartParams) {
                        editModel.steps[1].chartParams = {
                          categoryColIndex: 1,
                          seriesColIndices: [2],
                          chartType: 'column',
                          title: '图表分析结果',
                          targetCell: 'F2',
                          action: 'create_new',
                          errorHandling: 'reject_on_invalid'
                        };
                      }
                    }}
                  >
                    <option value="reconcile">两表主键对账 (reconcile)</option>
                    <option value="saved_macro">已保存宏 (saved_macro)</option>
                    <option value="chart">快捷图表生成 (chart)</option>
                  </select>
                </div>
                <div class="step-card-body">
                  {#if editModel.steps[1].toolType === 'reconcile' && editModel.steps[1].reconcileParams}
                    <div class="param-grid">
                      <div class="param-col">
                        <span class="form-label">左表 (自动接收第1步)</span>
                        <input type="text" class="input-text text-gray-500" value="消费前序输出新表" disabled />
                      </div>
                      <div class="param-col">
                        <span class="form-label">右表工作表</span>
                        <input type="text" class="input-text" bind:value={editModel.steps[1].reconcileParams.rightSheet} />
                      </div>
                      <div class="param-col">
                        <span class="form-label">右表区域</span>
                        <input type="text" class="input-text" bind:value={editModel.steps[1].reconcileParams.rightAddress} />
                      </div>
                    </div>
                  {:else if editModel.steps[1].toolType === 'saved_macro' && editModel.steps[1].macroParams}
                    <div class="param-grid">
                      <div class="param-col">
                        <span class="form-label">已保存宏 ID</span>
                        <input type="text" class="input-text" bind:value={editModel.steps[1].macroParams.scriptId} />
                      </div>
                      <div class="param-col">
                        <span class="form-label">入口过程</span>
                        <input type="text" class="input-text" bind:value={editModel.steps[1].macroParams.entryPoint} />
                      </div>
                    </div>
                  {:else if editModel.steps[1].toolType === 'chart' && editModel.steps[1].chartParams}
                    <div class="param-grid">
                      <div class="param-col">
                        <span class="form-label">数据输入源</span>
                        <input type="text" class="input-text text-gray-500" value="消费第1步确权输出 (prev_step_output)" disabled />
                      </div>
                      <div class="param-col">
                        <span class="form-label">类别 (X 轴) 列索引 (1-based)</span>
                        <input type="number" min="1" class="input-text" bind:value={editModel.steps[1].chartParams.categoryColIndex} />
                      </div>
                      <div class="param-col">
                        <span class="form-label">图表类型</span>
                        <select class="input-text" bind:value={editModel.steps[1].chartParams.chartType}>
                          <option value="column">柱状图 (Column)</option>
                          <option value="line">折线图 (Line)</option>
                          <option value="pie">饼图 (Pie)</option>
                        </select>
                      </div>
                      <div class="param-col">
                        <span class="form-label">图表标题</span>
                        <input type="text" class="input-text" bind:value={editModel.steps[1].chartParams.title} placeholder="图表标题" />
                      </div>
                      <div class="param-col">
                        <span class="form-label">放置起始单元格</span>
                        <input type="text" class="input-text" bind:value={editModel.steps[1].chartParams.targetCell} placeholder="例如: F2" />
                      </div>
                    </div>
                  {/if}
                </div>
              </div>

              <div class="edit-actions">
                <button class="btn btn-secondary" on:click={() => (isEditing = false)}>取消</button>
                <button class="btn btn-primary" on:click={handleSave}>
                  <Save size={13} />
                  <span>保存为新版本</span>
                </button>
              </div>
            </div>
          {:else if activeWorkflow}
            <!-- 运行与预检视图 -->
            <div class="workflow-dashboard">
              <div class="dashboard-header">
                <div class="title-area">
                  <h3>{activeWorkflow.name}</h3>
                  <p class="desc">{activeWorkflow.description || '暂无描述'}</p>
                </div>
                <div class="btn-group">
                  <button class="btn btn-secondary btn-sm" on:click={() => (isEditing = true)}>编辑参数</button>
                  {#if !activeWorkflow.isPreset}
                    <button class="btn btn-danger btn-sm" on:click={handleDelete} title="删除工作流">
                      <Trash2 size={13} />
                    </button>
                  {/if}
                </div>
              </div>

              <!-- 前置快照与恢复范围提示条 -->
              <div class="safety-banner">
                <ShieldAlert size={14} class="text-amber-600" />
                <div class="safety-text">
                  <strong>任务目标与快照恢复契约：</strong>
                  <span>任务执行严格绑定工作簿【<code>{workbook?.name || activeWorkflow.targetWorkbookName || '未指定'}</code>】。执行前强制创建整本物理副本。若失败，停止后续步骤并保留现场，用户可按已验证范围安全恢复。</span>
                </div>
              </div>

              <!-- 两步流程预览 -->
              <div class="pipeline-flow-view">
                <!-- 步骤 1 -->
                <div class="flow-step-box">
                  <div class="step-num">Step 1</div>
                  <div class="step-info">
                    <span class="step-title">{activeWorkflow.steps[0].stepName}</span>
                    <span class="step-type">工具类型: {activeWorkflow.steps[0].toolType}</span>
                  </div>
                </div>

                <div class="flow-arrow">
                  <ChevronRight size={16} />
                </div>

                <!-- 步骤 2 -->
                <div class="flow-step-box">
                  <div class="step-num">Step 2</div>
                  <div class="step-info">
                    <span class="step-title">{activeWorkflow.steps[1].stepName}</span>
                    <span class="step-type">工具类型: {activeWorkflow.steps[1].toolType}</span>
                  </div>
                </div>
              </div>

              <!-- 执行控制器 -->
              <div class="execute-bar">
                {#if isExecuting}
                  <button class="btn btn-danger" on:click={handleCancel}>
                    <StopCircle size={14} />
                    <span>在步骤边界取消</span>
                  </button>
                  <span class="running-indicator">正在受控调度 COM 串联执行，请稍候...</span>
                {:else}
                  <button
                    class="btn btn-success execute-btn"
                    disabled={!workbook || workbook.name === '未检测到活动工作簿'}
                    on:click={handleExecute}
                  >
                    <Play size={14} />
                    <span>开始执行流水线</span>
                  </button>
                  {#if !workbook || workbook.name === '未检测到活动工作簿'}
                    <span class="tip-warn">请先在 Excel 中打开目标工作簿</span>
                  {/if}
                {/if}
              </div>

              <!-- 当前运行实时结果卡片 -->
              {#if currentRun}
                <div class="run-result-card" class:card-success={currentRun.status === 'completed'} class:card-fail={currentRun.status.startsWith('stopped') || currentRun.status === 'blocked'}>
                  <div class="run-card-header">
                    <span class="run-status-tag status-{currentRun.status}">{currentRun.status}</span>
                    <span class="run-time">耗时: {currentRun.totalElapsedMs}ms</span>
                    {#if currentRun.snapshotId}
                      <span class="snapshot-tag">快照 ID: {currentRun.snapshotId}</span>
                    {/if}
                  </div>
                  <div class="run-summary-text">{currentRun.recoveryNotice || currentRun.failureMessage || '执行结束'}</div>

                  <!-- 步骤卡片展开 -->
                  <div class="steps-result-grid">
                    {#each currentRun.stepResults as sr}
                      <div class="step-result-item">
                        <div class="step-res-head">
                          <span>第 {sr.stepIndex} 步: {sr.stepName}</span>
                          <span class="res-status-badge status-{sr.status}">{sr.status}</span>
                        </div>
                        <div class="step-res-body">{sr.summaryText || sr.error || '无'}</div>
                        {#if sr.outputRef}
                          <div class="step-res-out">
                            产出工作表: <code>{sr.outputRef.sheetName}</code> (区域: {sr.outputRef.rangeAddress})
                          </div>
                        {/if}
                      </div>
                    {/each}
                  </div>
                </div>
              {/if}

              <!-- 历史运行记录 -->
              {#if runHistory.length > 0}
                <div class="history-section">
                  <h4>执行历史溯源 (最近 {runHistory.length} 次)</h4>
                  <div class="history-table">
                    {#each runHistory as rh}
                      <div class="history-row">
                        <span class="hist-time">{rh.executedAt.substring(11, 19)}</span>
                        <span class="hist-ver">v{rh.definitionVersion}</span>
                        <span class="hist-status status-{rh.status}">{rh.status}</span>
                        <span class="hist-steps">完成: {rh.completedSteps}/2 步</span>
                        <span class="hist-target">{rh.targetWorkbookName || '原表'}</span>
                      </div>
                    {/each}
                  </div>
                </div>
              {/if}
            </div>
          {:else}
            <div class="empty-state">
              <FileSpreadsheet size={32} class="text-gray-400" />
              <p>请在左侧选择已有流水线，或点击“新建”创建双步骤任务流水线。</p>
            </div>
          {/if}
        </div>
      </div>
    </div>
  </div>
{/if}

<style>
  .modal-backdrop {
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
  }

  .modal-window {
    background: #ffffff;
    width: 680px;
    max-width: 95vw;
    height: 600px;
    max-height: 90vh;
    border-radius: 6px;
    display: flex;
    flex-direction: column;
    box-shadow: 0 8px 30px rgba(0, 0, 0, 0.25);
    overflow: hidden;
  }

  .modal-header {
    height: 42px;
    background: #f8f9fa;
    border-bottom: 1px solid #e1dfdd;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 14px;
  }

  .header-title {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 13px;
    font-weight: 600;
    color: #323130;
  }

  .version-badge {
    background: #e1dfdd;
    color: #323130;
    font-size: 10px;
    padding: 1px 5px;
    border-radius: 3px;
  }

  .preset-tag {
    background: #e0f2fe;
    color: #0369a1;
    font-size: 10px;
    padding: 1px 5px;
    border-radius: 3px;
  }

  .close-btn {
    border: none;
    background: transparent;
    cursor: pointer;
    color: #605e5c;
    padding: 4px;
  }

  .modal-body-container {
    flex: 1;
    display: flex;
    overflow: hidden;
  }

  .sidebar-list {
    width: 200px;
    border-right: 1px solid #e1dfdd;
    background: #faf9f8;
    display: flex;
    flex-direction: column;
  }

  .sidebar-toolbar {
    height: 36px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 10px;
    border-bottom: 1px solid #edebe9;
  }

  .toolbar-label {
    font-size: 11px;
    font-weight: 600;
    color: #605e5c;
  }

  .add-btn {
    display: flex;
    align-items: center;
    gap: 3px;
    background: #ffffff;
    border: 1px solid #c8c6c4;
    padding: 2px 6px;
    border-radius: 3px;
    font-size: 10px;
    cursor: pointer;
  }

  .workflows-scroll {
    flex: 1;
    overflow-y: auto;
    padding: 6px;
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  .wf-item-btn {
    text-align: left;
    background: #ffffff;
    border: 1px solid #edebe9;
    border-radius: 4px;
    padding: 6px 8px;
    cursor: pointer;
    display: flex;
    flex-direction: column;
    gap: 2px;
  }

  .wf-item-btn.active {
    background: #eef2ff;
    border-color: #6366f1;
  }

  .wf-item-title {
    display: flex;
    justify-content: space-between;
    font-size: 11px;
    font-weight: 600;
    color: #323130;
  }

  .wf-item-ver {
    font-size: 9px;
    color: #8a8886;
  }

  .wf-item-desc {
    font-size: 10px;
    color: #605e5c;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .main-content-panel {
    flex: 1;
    overflow-y: auto;
    padding: 14px;
    display: flex;
    flex-direction: column;
    gap: 12px;
  }

  .alert-box {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 6px 10px;
    border-radius: 4px;
    font-size: 11px;
  }

  .error-alert {
    background: #fde7e9;
    color: #a80000;
    border: 1px solid #f8d7da;
  }

  .success-alert {
    background: #dff6dd;
    color: #107c41;
    border: 1px solid #c3e6cb;
  }

  .dashboard-header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
  }

  .dashboard-header h3 {
    margin: 0;
    font-size: 14px;
    color: #201f1e;
  }

  .dashboard-header .desc {
    margin: 3px 0 0 0;
    font-size: 11px;
    color: #605e5c;
  }

  .btn-group {
    display: flex;
    gap: 6px;
  }

  .safety-banner {
    display: flex;
    align-items: flex-start;
    gap: 8px;
    background: #fff8e5;
    border: 1px solid #ffeeba;
    padding: 8px 10px;
    border-radius: 4px;
    font-size: 11px;
    line-height: 1.4;
    color: #856404;
  }

  .pipeline-flow-view {
    display: flex;
    align-items: center;
    gap: 8px;
    background: #f8f9fa;
    padding: 10px;
    border-radius: 4px;
    border: 1px solid #edebe9;
  }

  .flow-step-box {
    flex: 1;
    background: #ffffff;
    border: 1px solid #d2d0ce;
    border-radius: 4px;
    padding: 8px;
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  .step-num {
    font-size: 10px;
    font-weight: 700;
    color: #6366f1;
  }

  .step-title {
    font-size: 11px;
    font-weight: 600;
    color: #323130;
  }

  .step-type {
    font-size: 10px;
    color: #8a8886;
  }

  .flow-arrow {
    color: #8a8886;
  }

  .execute-bar {
    display: flex;
    align-items: center;
    gap: 10px;
    padding-top: 4px;
  }

  .execute-btn {
    background: #107c41;
    color: white;
    border: none;
    padding: 6px 14px;
    border-radius: 4px;
    font-size: 12px;
    font-weight: 600;
    display: flex;
    align-items: center;
    gap: 6px;
    cursor: pointer;
  }

  .execute-btn:disabled {
    background: #c8c6c4;
    cursor: not-allowed;
  }

  .tip-warn {
    font-size: 11px;
    color: #d83b01;
  }

  .running-indicator {
    font-size: 11px;
    color: #6366f1;
    font-style: italic;
  }

  .run-result-card {
    border: 1px solid #e1dfdd;
    border-radius: 4px;
    padding: 10px;
    background: #ffffff;
    display: flex;
    flex-direction: column;
    gap: 8px;
  }

  .run-result-card.card-success {
    border-color: #107c41;
  }

  .run-result-card.card-fail {
    border-color: #a80000;
  }

  .run-card-header {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 11px;
  }

  .run-status-tag {
    font-size: 10px;
    font-weight: 600;
    padding: 1px 6px;
    border-radius: 3px;
    text-transform: uppercase;
  }

  .status-completed, .status-success {
    background: #dff6dd;
    color: #107c41;
  }

  .status-stopped_on_step1, .status-stopped_on_step2, .status-failed {
    background: #fde7e9;
    color: #a80000;
  }

  .status-skipped, .status-blocked, .status-cancelled {
    background: #fff4ce;
    color: #795e00;
  }

  .snapshot-tag {
    background: #edebe9;
    color: #323130;
    padding: 1px 5px;
    border-radius: 2px;
    font-family: monospace;
    font-size: 10px;
  }

  .run-summary-text {
    font-size: 11px;
    color: #323130;
  }

  .steps-result-grid {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }

  .step-result-item {
    background: #faf9f8;
    border: 1px solid #edebe9;
    border-radius: 3px;
    padding: 6px 8px;
    font-size: 10px;
  }

  .step-res-head {
    display: flex;
    justify-content: space-between;
    font-weight: 600;
    color: #323130;
    margin-bottom: 2px;
  }

  .history-section {
    display: flex;
    flex-direction: column;
    gap: 6px;
    border-top: 1px solid #edebe9;
    padding-top: 10px;
  }

  .history-section h4 {
    margin: 0;
    font-size: 11px;
    color: #605e5c;
  }

  .history-table {
    display: flex;
    flex-direction: column;
    gap: 3px;
  }

  .history-row {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 10px;
    background: #f8f9fa;
    padding: 4px 6px;
    border-radius: 2px;
  }

  .config-form {
    display: flex;
    flex-direction: column;
    gap: 10px;
  }

  .form-row {
    display: flex;
    flex-direction: column;
    gap: 3px;
  }

  .form-row .form-label {
    font-size: 11px;
    font-weight: 600;
    color: #323130;
  }

  .input-text {
    border: 1px solid #c8c6c4;
    border-radius: 3px;
    padding: 4px 6px;
    font-size: 11px;
  }

  .step-config-card {
    border: 1px solid #e1dfdd;
    border-radius: 4px;
    background: #ffffff;
    overflow: hidden;
  }

  .step-card-header {
    background: #f3f2f1;
    padding: 6px 8px;
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .step-badge {
    background: #6366f1;
    color: white;
    font-size: 9px;
    font-weight: 700;
    padding: 1px 5px;
    border-radius: 2px;
  }

  .input-step-name {
    flex: 1;
    border: 1px solid #c8c6c4;
    padding: 2px 5px;
    border-radius: 2px;
    font-size: 11px;
  }

  .select-tool {
    border: 1px solid #c8c6c4;
    padding: 2px 4px;
    border-radius: 2px;
    font-size: 10px;
  }

  .step-card-body {
    padding: 8px;
  }

  .param-grid {
    display: flex;
    gap: 8px;
  }

  .param-col {
    flex: 1;
    display: flex;
    flex-direction: column;
    gap: 2px;
  }

  .param-col .form-label {
    font-size: 10px;
    color: #605e5c;
  }

  .pipe-connector {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    padding: 2px 0;
  }

  .pipe-text {
    font-size: 10px;
    color: #6366f1;
    font-weight: 500;
  }

  .edit-actions {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    padding-top: 6px;
  }

  .btn {
    border: none;
    border-radius: 3px;
    padding: 4px 10px;
    font-size: 11px;
    cursor: pointer;
    display: flex;
    align-items: center;
    gap: 4px;
  }

  .btn-primary {
    background: #6366f1;
    color: white;
  }

  .btn-secondary {
    background: #edebe9;
    color: #323130;
  }

  .btn-danger {
    background: #d13438;
    color: white;
  }

  .btn-sm {
    padding: 3px 8px;
    font-size: 10px;
  }

  .empty-state {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 40px;
    color: #8a8886;
    font-size: 12px;
    gap: 8px;
  }

  @media (max-width: 480px) {
    .modal-body-container {
      flex-direction: column;
    }
    .sidebar-list {
      width: 100%;
      height: 110px;
      border-right: none;
      border-bottom: 1px solid #e1dfdd;
    }
    .pipeline-flow-view {
      flex-direction: column;
    }
    .flow-arrow {
      transform: rotate(90deg);
    }
    .modal-window {
      width: calc(100vw - 8px);
      height: calc(100vh - 12px);
    }
  }
</style>
