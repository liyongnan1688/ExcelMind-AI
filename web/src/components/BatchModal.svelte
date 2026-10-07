<script lang="ts">
  import { onMount, onDestroy } from 'svelte';
  import {
    X,
    FolderOpen,
    FileSpreadsheet,
    Play,
    StopCircle,
    CheckCircle2,
    XCircle,
    AlertTriangle,
    Clock,
    RotateCcw,
    Layers,
    Plus,
    Trash2,
    ShieldAlert,
    ExternalLink,
    Code,
    Check
  } from 'lucide-svelte';
  import {
    bridge,
    type BatchJobDefinition,
    type BatchJobSummary,
    type BatchFileTaskResult,
    type ScriptItem,
    type ScriptParameterDef,
    type VbaEntryPointInfo
  } from '../services/bridge';

  export let isOpen = false;
  export let onClose: () => void;

  const STORAGE_KEY_ACTIVE_JOB = 'excelmind_batch_active_job_id';

  // ==================== 界面主视图状态 ====================
  // 'config': 任务配置与参数固化视图
  // 'running': 执行与进度监控视图
  let viewMode: 'config' | 'running' = 'config';

  // ==================== 配置区状态 ====================
  let fileList: Array<{ path: string; name: string; dir: string }> = [];
  let manualFilePath = '';
  let outputDir = '';

  let availableScripts: ScriptItem[] = [];
  let selectedScriptId = '';
  let customMacroCode = '';
  let macroCodeToRun = '';
  let entryPoint = '';
  let candidateEntryPoints: VbaEntryPointInfo[] = [];
  let currentParameters: ScriptParameterDef[] = [];
  let parameterValues: Record<string, string> = {};

  // 固定策略：遇错即停
  const stopOnError = true;

  // 主动风险确认（默认未勾选）
  let riskAccepted = false;

  // 固化结果与防漂移状态
  let lockedJobDef: BatchJobDefinition | null = null;
  let validationError = '';
  let isLocking = false;

  // ==================== 执行区状态 ====================
  let activeJobId = '';
  let jobSummary: BatchJobSummary | null = null;
  let isStarting = false;
  let isCancelling = false;
  let cancelMessage = '';
  let restoreMessage = '';
  let restoreError = '';
  let openFolderMessage = '';
  let openFolderError = '';

  // 轮询控制（防重叠单并发机制）
  let pollingTimer: any = null;
  let isQueryInFlight = false;
  let isPollingActive = false;

  // 文件添加反馈提示
  let fileNotice = '';

  // ==================== 生命周期与状态恢复 ====================
  onMount(async () => {
    await loadSavedScripts();
    // 检查是否有跨重载未完成或最近的任务
    const savedJobId = localStorage.getItem(STORAGE_KEY_ACTIVE_JOB);
    if (savedJobId) {
      await attemptRestoreJob(savedJobId);
    }
  });

  onDestroy(() => {
    stopPolling();
  });

  $: if (isOpen) {
    onModalOpen();
  }

  async function onModalOpen() {
    restoreMessage = '';
    restoreError = '';
    const savedJobId = localStorage.getItem(STORAGE_KEY_ACTIVE_JOB);
    if (savedJobId) {
      await attemptRestoreJob(savedJobId);
    } else if (fileList.length === 0 && !outputDir) {
      // 默认尝试设定输出目录为用户的桌面或默认临时目录
      outputDir = '';
    }
  }

  async function loadSavedScripts() {
    try {
      const res = await bridge.getScripts();
      if (res.ok && res.data) {
        availableScripts = res.data;
      }
    } catch {
      // 忽略无法获取脚本库的错误
    }
  }

  // 仅恢复显示，绝不自动调用 start 重新执行
  async function attemptRestoreJob(jobId: string) {
    try {
      const res = await bridge.getBatchJobStatus(jobId);
      if (res.ok && res.data) {
        activeJobId = jobId;
        jobSummary = res.data;
        viewMode = 'running';
        restoreMessage = `已恢复任务显示: ${jobId}`;

        if (res.data.status === 'running') {
          startPolling(jobId);
        } else {
          stopPolling();
        }
      } else {
        restoreError = `无法恢复任务状态（${res.error || '任务不存在或后端已重启'}）`;
        localStorage.removeItem(STORAGE_KEY_ACTIVE_JOB);
      }
    } catch (e: any) {
      restoreError = `无法恢复任务状态: ${e.message || String(e)}`;
      localStorage.removeItem(STORAGE_KEY_ACTIVE_JOB);
    }
  }

  // ==================== 配置变更使固化失效 ====================
  function invalidateLock() {
    lockedJobDef = null;
    validationError = '';
    riskAccepted = false; // 配置变动后需重新确认风险
  }

  // ==================== 文件选择与去重逻辑 ====================
  async function handleBrowseFiles() {
    fileNotice = '';
    try {
      const res = await bridge.browseFiles('Excel 文件 (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls');
      if (res.ok && res.data && res.data.length > 0) {
        addFilesToList(res.data);
      }
    } catch (e: any) {
      fileNotice = '选择文件失败: ' + (e.message || String(e));
    }
  }

  function handleAddManualPath() {
    fileNotice = '';
    const raw = manualFilePath.trim().replace(/^["']|["']$/g, '');
    if (!raw) return;
    addFilesToList([raw]);
    manualFilePath = '';
  }

  function addFilesToList(paths: string[]) {
    let addedCount = 0;
    let duplicateCount = 0;

    for (const p of paths) {
      const trimmed = p.trim().replace(/^["']|["']$/g, '');
      if (!trimmed) continue;

      // 检查路径完全一致去重
      const exists = fileList.some((f) => f.path.toLowerCase() === trimmed.toLowerCase());
      if (exists) {
        duplicateCount++;
        continue;
      }

      // 拆分文件名与父目录
      const lastSlash = Math.max(trimmed.lastIndexOf('\\'), trimmed.lastIndexOf('/'));
      const name = lastSlash >= 0 ? trimmed.substring(lastSlash + 1) : trimmed;
      const dir = lastSlash >= 0 ? trimmed.substring(0, lastSlash) : '';

      fileList = [...fileList, { path: trimmed, name, dir }];
      addedCount++;
    }

    if (addedCount > 0) {
      invalidateLock();
    }

    if (duplicateCount > 0) {
      fileNotice = `已添加 ${addedCount} 个文件，自动忽略 ${duplicateCount} 个同路径重复文件。`;
    } else {
      fileNotice = `已添加 ${addedCount} 个文件。`;
    }
  }

  function handleRemoveFile(index: number) {
    fileList = fileList.filter((_, i) => i !== index);
    invalidateLock();
  }

  function handleClearFiles() {
    fileList = [];
    invalidateLock();
  }

  async function handleBrowseOutputDir() {
    try {
      const res = await bridge.browseFolder('请选择批量处理输出目录');
      if (res.ok && res.data) {
        outputDir = res.data;
        invalidateLock();
      }
    } catch {
      // 忽略
    }
  }

  // ==================== 宏选择与参数提取 ====================
  function handleSelectScript(e: Event) {
    const id = (e.target as HTMLSelectElement).value;
    selectedScriptId = id;
    if (!id) {
      customMacroCode = '';
      macroCodeToRun = '';
      entryPoint = '';
      candidateEntryPoints = [];
      currentParameters = [];
      parameterValues = {};
      invalidateLock();
      return;
    }

    const script = availableScripts.find((s) => s.id === id);
    if (script) {
      customMacroCode = script.code;
      macroCodeToRun = script.code;
      candidateEntryPoints = script.candidateEntryPoints || [];
      if (candidateEntryPoints.length > 0) {
        entryPoint = candidateEntryPoints[0].name;
        currentParameters = candidateEntryPoints[0].parameters || [];
      } else {
        entryPoint = script.entryPoint || '';
        currentParameters = script.parameters || [];
      }
      parameterValues = {};
      currentParameters.forEach((p) => {
        parameterValues[p.name] = p.defaultValue || '';
      });
      invalidateLock();
    }
  }

  function handleEntryPointChange(e: Event) {
    const name = (e.target as HTMLSelectElement).value;
    entryPoint = name;
    const ep = candidateEntryPoints.find((x) => x.name === name);
    if (ep) {
      currentParameters = ep.parameters || [];
      parameterValues = {};
      currentParameters.forEach((p) => {
        parameterValues[p.name] = p.defaultValue || '';
      });
    }
    invalidateLock();
  }

  function handleCodeInput() {
    macroCodeToRun = customMacroCode;
    invalidateLock();
  }

  // ==================== 固化与启动流程 ====================
  async function handleValidateAndLock() {
    validationError = '';
    isLocking = true;

    if (fileList.length === 0) {
      validationError = '请至少添加一个待处理的 Excel 工作簿。';
      isLocking = false;
      return;
    }
    if (!macroCodeToRun.trim()) {
      validationError = '请输入或选择要批量执行的 VBA 宏代码。';
      isLocking = false;
      return;
    }
    if (!outputDir.trim()) {
      validationError = '请指定批量处理输出文件的保存目录。';
      isLocking = false;
      return;
    }

    try {
      const paths = fileList.map((f) => f.path);
      const paramsJson = currentParameters.length > 0 ? JSON.stringify(parameterValues) : '';

      const res = await bridge.validateBatchJob({
        filePaths: paths,
        code: macroCodeToRun,
        entryPoint: entryPoint || undefined,
        parameters: paramsJson || undefined,
        outputDir: outputDir.trim(),
        stopOnError: true
      });

      if (res.ok && res.data) {
        lockedJobDef = res.data;
      } else {
        validationError = res.error || '固化校验失败';
      }
    } catch (e: any) {
      validationError = '固化校验异常: ' + (e.message || String(e));
    } finally {
      isLocking = false;
    }
  }

  async function handleStartBatch() {
    if (!lockedJobDef) {
      await handleValidateAndLock();
      if (!lockedJobDef) return;
    }

    if (!riskAccepted) {
      validationError = '请先主动确认阅读并知晓批量执行风险声明。';
      return;
    }

    isStarting = true;
    validationError = '';
    cancelMessage = '';

    try {
      // 启动优先只传 jobId；附带其他字段必须 100% 一致
      const res = await bridge.startBatchJob({
        jobId: lockedJobDef.jobId,
        filePaths: lockedJobDef.files.map((f) => f.originalFilePath),
        code: lockedJobDef.macroCode,
        entryPoint: lockedJobDef.entryPoint,
        parameters: lockedJobDef.parametersJson,
        outputDir: lockedJobDef.outputDir,
        stopOnError: true,
        sync: false
      });

      if (res.ok && res.data) {
        activeJobId = lockedJobDef.jobId;
        jobSummary = res.data;
        localStorage.setItem(STORAGE_KEY_ACTIVE_JOB, activeJobId);
        viewMode = 'running';
        startPolling(activeJobId);
      } else {
        validationError = res.error || '启动批量任务失败';
      }
    } catch (e: any) {
      validationError = '启动批量任务异常: ' + (e.message || String(e));
    } finally {
      isStarting = false;
    }
  }

  // ==================== 防重叠单并发轮询机制 ====================
  function startPolling(jobId: string) {
    stopPolling();
    isPollingActive = true;

    const poll = async () => {
      if (!isPollingActive || !isOpen) return;
      if (isQueryInFlight) {
        // 上一次请求尚未返回，等待下一次循环，彻底避免请求重叠与并发竞争
        pollingTimer = setTimeout(poll, 600);
        return;
      }

      isQueryInFlight = true;
      try {
        const res = await bridge.getBatchJobStatus(jobId);
        if (res.ok && res.data) {
          jobSummary = res.data;
          // 若达到终端态，停止轮询
          if (
            res.data.status === 'completed' ||
            res.data.status === 'stopped_on_error' ||
            res.data.status === 'cancelled'
          ) {
            stopPolling();
            return;
          }
        } else {
          // 查询出错或任务失效，停止轮询
          stopPolling();
          return;
        }
      } catch {
        // 网络/桥接抖动，继续等待下一轮
      } finally {
        isQueryInFlight = false;
      }

      if (isPollingActive) {
        pollingTimer = setTimeout(poll, 800);
      }
    };

    pollingTimer = setTimeout(poll, 200);
  }

  function stopPolling() {
    isPollingActive = false;
    if (pollingTimer) {
      clearTimeout(pollingTimer);
      pollingTimer = null;
    }
  }

  // ==================== 取消与打开输出目录 ====================
  async function handleCancel() {
    if (!activeJobId) return;
    isCancelling = true;
    try {
      const res = await bridge.cancelBatchJob(activeJobId);
      if (res.ok) {
        cancelMessage = res.message || '已登记取消请求。当前正在处理的文件结束后将停止后续。';
      }
    } catch (e: any) {
      cancelMessage = '发送取消请求失败: ' + (e.message || String(e));
    } finally {
      isCancelling = false;
    }
  }

  async function handleOpenOutputFolder() {
    openFolderMessage = '';
    openFolderError = '';
    try {
      const dir = jobSummary?.fileResults?.find((r) => r.finalOutputPath)?.finalOutputPath || outputDir;
      const res = await bridge.openOutputFolder({
        outputDir: outputDir || dir,
        jobId: activeJobId
      });
      if (res.ok) {
        openFolderMessage = res.message || '已打开输出目录';
      } else {
        openFolderError = res.error || '无法打开输出目录';
      }
    } catch (e: any) {
      openFolderError = '打开失败: ' + (e.message || String(e));
    }
  }

  function handleResetConfig() {
    stopPolling();
    viewMode = 'config';
    activeJobId = '';
    jobSummary = null;
    invalidateLock();
    localStorage.removeItem(STORAGE_KEY_ACTIVE_JOB);
  }

  function handleCloseModal() {
    // 关闭面板不取消后台任务，只停止前端定时轮询，保留任务 ID 供重开恢复
    stopPolling();
    onClose();
  }
</script>

{#if isOpen}
  <div
    class="batch-modal-backdrop"
    role="presentation"
    on:click|self={handleCloseModal}
    on:keydown={(e) => e.key === 'Escape' && handleCloseModal()}
  >
    <div class="batch-modal-window" role="dialog" aria-modal="true" aria-labelledby="batch-modal-title">
      <!-- 模态框顶部 -->
      <div class="modal-header">
        <div class="header-title-box">
          <Layers size={18} class="text-emerald-700" />
          <h2 id="batch-modal-title">批量宏处理面板 (TASK-R4b)</h2>
          <span class="badge-stage">隔离工作副本 · 遇错即停</span>
        </div>
        <button class="close-btn" on:click={handleCloseModal} aria-label="关闭面板">
          <X size={16} />
        </button>
      </div>

      <!-- 提示栏：面板关闭不取消任务 -->
      <div class="header-notice-bar">
        <span>💡 提示：面板关闭不会中断后台批量任务。重新打开将自动恢复显示最新进度。</span>
      </div>

      <!-- 主体内容 -->
      <div class="modal-body">
        {#if restoreError}
          <div class="error-banner">
            <AlertTriangle size={15} />
            <span>{restoreError}</span>
            <button class="retry-link" on:click={() => (restoreError = '')}>关闭提示</button>
          </div>
        {/if}

        {#if viewMode === 'config'}
          <!-- ==================== 配置视图 ==================== -->
          <div class="config-container">
            <!-- 1. 待处理文件选择与清单 -->
            <section class="config-section">
              <div class="section-title-row">
                <div class="title-with-count">
                  <span class="step-num">1</span>
                  <h3>选择待处理 Excel 工作簿</h3>
                  <span class="count-tag">已选 {fileList.length} 个文件</span>
                </div>
                <div class="action-buttons-group">
                  <button class="btn-subtle" on:click={handleBrowseFiles} type="button">
                    <Plus size={14} /> 添加文件 (浏览)
                  </button>
                  {#if fileList.length > 0}
                    <button class="btn-danger-text" on:click={handleClearFiles} type="button">
                      <Trash2 size={13} /> 清空
                    </button>
                  {/if}
                </div>
              </div>

              <!-- 手动路径添加 -->
              <div class="manual-input-row">
                <input
                  type="text"
                  placeholder="或粘贴单个绝对文件路径，按回车添加..."
                  bind:value={manualFilePath}
                  on:keydown={(e) => e.key === 'Enter' && handleAddManualPath()}
                />
                <button class="btn-secondary" type="button" on:click={handleAddManualPath}>添加</button>
              </div>

              {#if fileNotice}
                <div class="file-notice-text">{fileNotice}</div>
              {/if}

              <!-- 文件列表容器 -->
              <div class="file-list-box">
                {#if fileList.length === 0}
                  <div class="empty-list-placeholder">
                    <FileSpreadsheet size={28} class="text-gray-300 mb-1" />
                    <span>暂未添加文件。支持格式：.xlsx, .xlsm, .xls</span>
                  </div>
                {:else}
                  <div class="file-table-wrapper">
                    {#each fileList as file, idx}
                      <div class="file-row-item">
                        <span class="file-index">{idx + 1}</span>
                        <div class="file-info-cell">
                          <span class="file-name-text" title={file.name}>{file.name}</span>
                          <span class="file-dir-text" title={file.path}>{file.dir || file.path}</span>
                        </div>
                        <button
                          class="file-remove-btn"
                          on:click={() => handleRemoveFile(idx)}
                          title="移除此文件"
                          type="button"
                        >
                          <X size={13} />
                        </button>
                      </div>
                    {/each}
                  </div>
                {/if}
              </div>
            </section>

            <!-- 2. 宏选择、入口与结构化参数 -->
            <section class="config-section">
              <div class="section-title-row">
                <div class="title-with-count">
                  <span class="step-num">2</span>
                  <h3>选择执行宏与过程入口</h3>
                </div>
              </div>

              <div class="form-grid-two">
                <!-- 预设宏库下拉 -->
                <div class="form-item">
                  <label for="script-select">从已有宏库选择：</label>
                  <select id="script-select" value={selectedScriptId} on:change={handleSelectScript}>
                    <option value="">-- 手动输入或自定义代码 --</option>
                    {#each availableScripts as s}
                      <option value={s.id}>{s.displayName || s.name}</option>
                    {/each}
                  </select>
                </div>

                <!-- 过程入口下拉 -->
                <div class="form-item">
                  <label for="entry-point-select">执行过程 (Sub)：</label>
                  {#if candidateEntryPoints.length > 0}
                    <select id="entry-point-select" value={entryPoint} on:change={handleEntryPointChange}>
                      {#each candidateEntryPoints as ep}
                        <option value={ep.name}>{ep.name} ({ep.parameters.length} 参)</option>
                      {/each}
                    </select>
                  {:else}
                    <input
                      id="entry-point-select"
                      type="text"
                      placeholder="过程名 (例如: ProcessAllData)"
                      bind:value={entryPoint}
                      on:input={invalidateLock}
                    />
                  {/if}
                </div>
              </div>

              <!-- VBA 源码预览/编辑 -->
              <div class="form-item mt-2">
                <label for="macro-code-textarea">VBA 源码：</label>
                <textarea
                  id="macro-code-textarea"
                  class="code-textarea"
                  placeholder="Sub ProcessCurrent()\n    ActiveSheet.Range('A1').Value = 'Processed'\nEnd Sub"
                  bind:value={customMacroCode}
                  on:input={handleCodeInput}
                  rows={4}
                ></textarea>
              </div>

              <!-- 结构化参数表单 (复用 R2c 契约) -->
              {#if currentParameters.length > 0}
                <div class="param-binding-card mt-2">
                  <div class="param-header">
                    <span class="font-medium">过程参数配置 (将逐文件在隔离副本中求值绑定)：</span>
                  </div>
                  <div class="param-fields-grid">
                    {#each currentParameters as param}
                      <div class="param-field-item">
                        <label for="param-{param.name}">
                          {param.name} <span class="param-type">({param.type})</span>:
                        </label>
                        <input
                          id="param-{param.name}"
                          type="text"
                          placeholder={param.defaultValue || `请输入 ${param.name}`}
                          bind:value={parameterValues[param.name]}
                          on:input={invalidateLock}
                        />
                      </div>
                    {/each}
                  </div>
                  <p class="param-tip">
                    注：Worksheet / Range 参数将在各文件的独立隔离副本中自动绑定；若文件缺少对应表名或区域将在流水线预检阶段阻断。
                  </p>
                </div>
              {/if}
            </section>

            <!-- 3. 输出目录 -->
            <section class="config-section">
              <div class="section-title-row">
                <div class="title-with-count">
                  <span class="step-num">3</span>
                  <h3>确认输出目录</h3>
                </div>
                <button class="btn-subtle" on:click={handleBrowseOutputDir} type="button">
                  <FolderOpen size={14} /> 浏览目录
                </button>
              </div>
              <div class="manual-input-row">
                <input
                  type="text"
                  placeholder="输出文件夹绝对路径 (例如: C:\ExcelOutputs)"
                  bind:value={outputDir}
                  on:input={invalidateLock}
                />
              </div>
              <p class="section-tip">
                所有成功处理的文件将保存在此目录下，同名文件将自动递增编号防覆盖。原文件绝对保持逐字节恒定。
              </p>
            </section>

            <!-- 4. 执行策略与前置风险声明 -->
            <section class="config-section risk-section">
              <div class="section-title-row">
                <div class="title-with-count">
                  <span class="step-num">4</span>
                  <h3 class="text-amber-800">执行策略与安全风险声明</h3>
                </div>
                <span class="policy-pill">执行策略：遇错即停 (Stop on Error，固定)</span>
              </div>

              <!-- 风险声明明文展示 -->
              <div class="risk-notice-card">
                <div class="risk-card-header">
                  <ShieldAlert size={16} class="text-amber-600" />
                  <strong>重要安全声明（请务必在启动前充分阅读）：</strong>
                </div>
                <ul class="risk-list">
                  <li>
                    <strong>原文件保护：</strong>
                    本插件使用专用受控工作副本执行宏，保证原文件与当前 Excel 中已打开的工作簿不被修改或覆盖。
                  </li>
                  <li>
                    <strong>非安全沙箱：</strong>
                    工作副本机制绝不是任意 VBA 的安全沙箱！若宏代码中包含固定绝对文件路径、外部文件写入/删除 (如 Kill,
                    FileSystemObject) 或系统命令，仍会直接作用于外部环境。
                  </li>
                  <li>
                    <strong>责任提示：</strong>
                    严禁对未知、不可信或包含外部系统副作用的宏进行批量执行。
                  </li>
                </ul>

                <!-- 主动确认复选框（默认未勾选） -->
                <div class="risk-confirm-row">
                  <label class="checkbox-label" for="risk-confirm-checkbox">
                    <input
                      id="risk-confirm-checkbox"
                      type="checkbox"
                      bind:checked={riskAccepted}
                    />
                    <span class="checkbox-text">
                      我已充分阅读并理解上述风险声明，确认知晓批量执行的不可逆外部影响，同意启动执行。
                    </span>
                  </label>
                </div>
              </div>

              {#if validationError}
                <div class="validation-error-bar mt-2">
                  <AlertTriangle size={15} />
                  <span>{validationError}</span>
                </div>
              {/if}

              <!-- 固化与启动按钮 -->
              <div class="action-footer-row mt-3">
                <div class="lock-status-indicator">
                  {#if lockedJobDef}
                    <span class="locked-badge">
                      <Check size={13} /> 队列已固化 (ID: {lockedJobDef.jobId})
                    </span>
                  {:else}
                    <span class="unlocked-text">待固化校验</span>
                  {/if}
                </div>

                <div class="footer-buttons">
                  {#if !lockedJobDef}
                    <button
                      class="btn-primary"
                      type="button"
                      disabled={isLocking || fileList.length === 0}
                      on:click={handleValidateAndLock}
                    >
                      {isLocking ? '正在预检固化...' : '预检并固化队列'}
                    </button>
                  {:else}
                    <button
                      class="btn-success"
                      type="button"
                      disabled={!riskAccepted || isStarting}
                      on:click={handleStartBatch}
                    >
                      <Play size={14} />
                      {isStarting ? '正在启动...' : '开始批量执行'}
                    </button>
                  {/if}
                </div>
              </div>
            </section>
          </div>
        {:else}
          <!-- ==================== 执行与进度监控视图 ==================== -->
          <div class="running-container">
            <!-- 运行状态横幅 -->
            <div class="status-summary-card">
              <div class="status-main-row">
                <div class="status-badge-title">
                  {#if jobSummary?.status === 'running'}
                    <span class="status-pill pill-running">
                      <Clock size={14} class="spin" /> 正在串行执行批量任务
                    </span>
                  {:else if jobSummary?.status === 'completed'}
                    <span class="status-pill pill-completed">
                      <CheckCircle2 size={14} /> 批量任务已结束
                    </span>
                  {:else if jobSummary?.status === 'stopped_on_error'}
                    <span class="status-pill pill-error">
                      <AlertTriangle size={14} /> 任务因遇到错误已中止（未全部成功）
                    </span>
                  {:else if jobSummary?.status === 'cancelled'}
                    <span class="status-pill pill-cancelled">
                      <StopCircle size={14} /> 任务已取消
                    </span>
                  {/if}
                  <span class="job-id-text">任务 ID: {activeJobId}</span>
                </div>

                <div class="status-actions">
                  {#if jobSummary?.status === 'running'}
                    <button
                      class="btn-danger-outline"
                      type="button"
                      disabled={isCancelling}
                      on:click={handleCancel}
                    >
                      <StopCircle size={14} />
                      {isCancelling ? '正在取消...' : '取消任务'}
                    </button>
                  {:else}
                    <button class="btn-secondary" type="button" on:click={handleResetConfig}>
                      <RotateCcw size={14} /> 新建批量任务
                    </button>
                  {/if}
                  <button class="btn-primary" type="button" on:click={handleOpenOutputFolder}>
                    <ExternalLink size={14} /> 打开输出目录
                  </button>
                </div>
              </div>

              {#if cancelMessage}
                <div class="cancel-notice mt-2">{cancelMessage}</div>
              {/if}
              {#if openFolderMessage}
                <div class="folder-notice mt-2 text-emerald-700">{openFolderMessage}</div>
              {/if}
              {#if openFolderError}
                <div class="folder-notice mt-2 text-rose-600">{openFolderError}</div>
              {/if}

              <!-- 客观指标统计卡片 (6 项真实指标) -->
              <div class="metrics-grid mt-3">
                <div class="metric-card">
                  <span class="metric-num">{jobSummary?.totalFiles || 0}</span>
                  <span class="metric-label">总文件数</span>
                </div>
                <div class="metric-card">
                  <span class="metric-num">
                    {(jobSummary?.successCount || 0) +
                      (jobSummary?.failedCount || 0) +
                      (jobSummary?.blockedCount || 0) +
                      (jobSummary?.cancelledCount || 0)}
                  </span>
                  <span class="metric-label">已完成数</span>
                </div>
                <div class="metric-card card-success">
                  <span class="metric-num">{jobSummary?.successCount || 0}</span>
                  <span class="metric-label">成功</span>
                </div>
                <div class="metric-card card-failed">
                  <span class="metric-num">{jobSummary?.failedCount || 0}</span>
                  <span class="metric-label">失败</span>
                </div>
                <div class="metric-card card-blocked">
                  <span class="metric-num">{jobSummary?.blockedCount || 0}</span>
                  <span class="metric-label">阻断</span>
                </div>
                <div class="metric-card card-cancelled">
                  <span class="metric-num">{jobSummary?.cancelledCount || 0}</span>
                  <span class="metric-label">已取消</span>
                </div>
                <div class="metric-card card-pending">
                  <span class="metric-num">{jobSummary?.pendingCount || 0}</span>
                  <span class="metric-label">未执行</span>
                </div>
              </div>
            </div>

            <!-- 逐文件明细流水线 -->
            <div class="results-table-section mt-3">
              <div class="table-header-title">
                <h3>逐文件处理明细</h3>
                <span class="table-hint">
                  {jobSummary?.status === 'running' ? '后台受控 STA 线程实时更新' : '全部执行记录'}
                </span>
              </div>

              <div class="results-list-box">
                {#if !jobSummary?.fileResults || jobSummary.fileResults.length === 0}
                  <div class="empty-list-placeholder">暂无文件明细数据</div>
                {:else}
                  {#each jobSummary.fileResults as res}
                    <div class="result-row-card status-{res.status}">
                      <div class="row-main-line">
                        <div class="row-file-badge">
                          <span class="file-num">#{res.fileIndex + 1}</span>
                          <span class="result-file-name" title={res.originalFilePath}>{res.originalFileName}</span>
                        </div>

                        <div class="row-status-badge">
                          {#if res.status === 'success'}
                            <span class="status-tag tag-success">
                              <CheckCircle2 size={13} /> 成功 ({res.elapsedMs}ms)
                            </span>
                          {:else if res.status === 'failed'}
                            <span class="status-tag tag-failed">
                              <XCircle size={13} /> 失败 (阶段: {res.failureStage || '未知'})
                            </span>
                          {:else if res.status === 'blocked'}
                            <span class="status-tag tag-blocked">
                              <AlertTriangle size={13} /> 阻断 (阶段: {res.failureStage || '校验'})
                            </span>
                          {:else if res.status === 'running'}
                            <span class="status-tag tag-running">
                              <Clock size={13} class="spin" /> 执行中...
                            </span>
                          {:else if res.status === 'cancelled'}
                            <span class="status-tag tag-cancelled">
                              <StopCircle size={13} /> 已取消
                            </span>
                          {:else}
                            <span class="status-tag tag-pending">
                              <Clock size={13} /> 等待中
                            </span>
                          {/if}
                        </div>
                      </div>

                      <!-- 附加详细说明 (输出路径或错误现场) -->
                      {#if res.finalOutputPath}
                        <div class="result-sub-line">
                          <span class="sub-label">输出:</span>
                          <span class="sub-content" title={res.finalOutputPath}>{res.finalOutputPath}</span>
                        </div>
                      {/if}

                      {#if res.error}
                        <div class="result-sub-line error-line">
                          <span class="sub-label">错误:</span>
                          <span class="sub-content">{res.error}</span>
                        </div>
                      {/if}

                      {#if res.workingCopyPath && (res.status === 'failed' || res.status === 'blocked')}
                        <div class="result-sub-line tip-line">
                          <span class="sub-label">现场副本保留在:</span>
                          <span class="sub-content" title={res.workingCopyPath}>{res.workingCopyPath}</span>
                        </div>
                      {/if}
                    </div>
                  {/each}
                {/if}
              </div>
            </div>
          </div>
        {/if}
      </div>

      <!-- 模态框底部固定栏 -->
      <div class="modal-footer">
        <button class="btn-secondary" type="button" on:click={handleCloseModal}>关闭面板</button>
      </div>
    </div>
  </div>
{/if}

<style>
  .batch-modal-backdrop {
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background: rgba(0, 0, 0, 0.45);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 1050;
    backdrop-filter: blur(2px);
  }

  .batch-modal-window {
    width: 92vw;
    max-width: 820px;
    height: 90vh;
    max-height: 820px;
    background: #ffffff;
    border-radius: 8px;
    box-shadow: 0 12px 36px rgba(0, 0, 0, 0.22);
    display: flex;
    flex-direction: column;
    overflow: hidden;
    font-size: 12px;
    color: #323130;
  }

  .modal-header {
    height: 46px;
    background: #fcfcfc;
    border-bottom: 1px solid #edebe9;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 16px;
    flex-shrink: 0;
  }

  .header-title-box {
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .header-title-box h2 {
    margin: 0;
    font-size: 14px;
    font-weight: 600;
    color: #201f1e;
  }

  .badge-stage {
    font-size: 10px;
    background: #e1dfdd;
    color: #323130;
    padding: 2px 6px;
    border-radius: 4px;
  }

  .close-btn {
    border: none;
    background: transparent;
    cursor: pointer;
    color: #605e5c;
    padding: 4px;
    border-radius: 4px;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .close-btn:hover {
    background: #f3f2f1;
    color: #201f1e;
  }

  .header-notice-bar {
    background: #f0f7f3;
    border-bottom: 1px solid #d4ebdd;
    color: #107c41;
    font-size: 11px;
    padding: 6px 16px;
    flex-shrink: 0;
  }

  .modal-body {
    flex: 1;
    overflow-y: auto;
    padding: 16px;
    background: #faf9f8;
  }

  .error-banner {
    background: #fdf3f2;
    border: 1px solid #f8d7da;
    color: #a80000;
    padding: 8px 12px;
    border-radius: 4px;
    display: flex;
    align-items: center;
    gap: 8px;
    margin-bottom: 12px;
  }

  .retry-link {
    background: none;
    border: none;
    color: #a80000;
    text-decoration: underline;
    cursor: pointer;
    font-size: 11px;
    margin-left: auto;
  }

  /* 配置区域布局 */
  .config-container {
    display: flex;
    flex-direction: column;
    gap: 16px;
  }

  .config-section {
    background: #ffffff;
    border: 1px solid #edebe9;
    border-radius: 6px;
    padding: 14px;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
  }

  .section-title-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 10px;
  }

  .title-with-count {
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .step-num {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 20px;
    height: 20px;
    background: #107c41;
    color: #ffffff;
    border-radius: 50%;
    font-size: 11px;
    font-weight: 600;
  }

  .title-with-count h3 {
    margin: 0;
    font-size: 13px;
    font-weight: 600;
    color: #201f1e;
  }

  .count-tag {
    font-size: 11px;
    color: #605e5c;
    background: #f3f2f1;
    padding: 2px 6px;
    border-radius: 10px;
  }

  .action-buttons-group {
    display: flex;
    gap: 8px;
  }

  .btn-subtle {
    background: #f3f2f1;
    border: 1px solid #d2d0ce;
    color: #201f1e;
    padding: 4px 8px;
    border-radius: 4px;
    cursor: pointer;
    display: flex;
    align-items: center;
    gap: 4px;
    font-size: 11px;
  }
  .btn-subtle:hover {
    background: #edebe9;
  }

  .btn-danger-text {
    background: transparent;
    border: none;
    color: #a80000;
    padding: 4px 8px;
    border-radius: 4px;
    cursor: pointer;
    display: flex;
    align-items: center;
    gap: 4px;
    font-size: 11px;
  }
  .btn-danger-text:hover {
    background: #fdf3f2;
  }

  .manual-input-row {
    display: flex;
    gap: 8px;
    margin-bottom: 8px;
  }

  .manual-input-row input {
    flex: 1;
    height: 30px;
    padding: 0 8px;
    border: 1px solid #c8c6c4;
    border-radius: 4px;
    font-size: 11px;
  }

  .file-notice-text {
    font-size: 11px;
    color: #107c41;
    margin-bottom: 6px;
  }

  .file-list-box {
    border: 1px solid #edebe9;
    border-radius: 4px;
    background: #faf9f8;
    max-height: 140px;
    overflow-y: auto;
  }

  .empty-list-placeholder {
    padding: 20px;
    text-align: center;
    color: #a19f9d;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
  }

  .file-table-wrapper {
    display: flex;
    flex-direction: column;
  }

  .file-row-item {
    display: flex;
    align-items: center;
    padding: 6px 10px;
    border-bottom: 1px solid #f3f2f1;
    background: #ffffff;
  }
  .file-row-item:last-child {
    border-bottom: none;
  }

  .file-index {
    width: 24px;
    color: #8a8886;
    font-size: 10px;
    flex-shrink: 0;
  }

  .file-info-cell {
    flex: 1;
    min-width: 0;
    display: flex;
    flex-direction: column;
  }

  .file-name-text {
    font-weight: 500;
    color: #201f1e;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .file-dir-text {
    font-size: 10px;
    color: #605e5c;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .file-remove-btn {
    border: none;
    background: transparent;
    cursor: pointer;
    color: #a19f9d;
    padding: 4px;
    border-radius: 3px;
  }
  .file-remove-btn:hover {
    color: #a80000;
    background: #fdf3f2;
  }

  /* 宏表单 */
  .form-grid-two {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
  }

  .form-item {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  .form-item label {
    font-size: 11px;
    color: #605e5c;
  }

  .form-item select,
  .form-item input {
    height: 30px;
    border: 1px solid #c8c6c4;
    border-radius: 4px;
    padding: 0 8px;
    font-size: 11px;
    background: #ffffff;
  }

  .code-textarea {
    width: 100%;
    border: 1px solid #c8c6c4;
    border-radius: 4px;
    padding: 8px;
    font-family: Consolas, monospace;
    font-size: 11px;
    background: #fcfcfc;
    resize: vertical;
  }

  /* 参数卡片 */
  .param-binding-card {
    background: #f3f9f5;
    border: 1px solid #cbe5d5;
    border-radius: 4px;
    padding: 10px;
  }

  .param-header {
    font-size: 11px;
    color: #107c41;
    margin-bottom: 8px;
  }

  .param-fields-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
    gap: 8px;
  }

  .param-field-item {
    display: flex;
    flex-direction: column;
    gap: 2px;
  }

  .param-field-item label {
    font-size: 10px;
    color: #323130;
  }

  .param-type {
    color: #605e5c;
    font-weight: normal;
  }

  .param-field-item input {
    height: 26px;
    border: 1px solid #b8dfc7;
    border-radius: 3px;
    padding: 0 6px;
    font-size: 11px;
  }

  .param-tip {
    font-size: 10px;
    color: #605e5c;
    margin: 6px 0 0 0;
  }

  .section-tip {
    font-size: 10px;
    color: #605e5c;
    margin: 4px 0 0 0;
  }

  /* 安全与风险声明 */
  .risk-section {
    border-color: #f7e1b5;
    background: #fffdfa;
  }

  .policy-pill {
    font-size: 10px;
    background: #fff4ce;
    color: #795b00;
    border: 1px solid #fce28e;
    padding: 2px 8px;
    border-radius: 12px;
    font-weight: 500;
  }

  .risk-notice-card {
    background: #fffbf0;
    border: 1px solid #fae2a6;
    border-radius: 4px;
    padding: 12px;
  }

  .risk-card-header {
    display: flex;
    align-items: center;
    gap: 6px;
    color: #8f6200;
    font-size: 11px;
    margin-bottom: 6px;
  }

  .risk-list {
    margin: 0;
    padding-left: 18px;
    color: #494846;
    font-size: 11px;
    line-height: 1.6;
  }

  .risk-confirm-row {
    margin-top: 10px;
    padding-top: 8px;
    border-top: 1px dashed #fae2a6;
  }

  .checkbox-label {
    display: flex;
    align-items: flex-start;
    gap: 8px;
    cursor: pointer;
  }

  .checkbox-label input {
    margin-top: 2px;
    cursor: pointer;
  }

  .checkbox-text {
    font-size: 11px;
    font-weight: 600;
    color: #795b00;
    line-height: 1.4;
  }

  .validation-error-bar {
    background: #fdf3f2;
    border: 1px solid #f8d7da;
    color: #a80000;
    padding: 6px 10px;
    border-radius: 4px;
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 11px;
  }

  .action-footer-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
  }

  .locked-badge {
    background: #dff6dd;
    color: #107c41;
    border: 1px solid #a3dfa0;
    padding: 3px 8px;
    border-radius: 4px;
    font-size: 11px;
    display: inline-flex;
    align-items: center;
    gap: 4px;
  }

  .unlocked-text {
    color: #a19f9d;
    font-size: 11px;
  }

  .footer-buttons {
    display: flex;
    gap: 10px;
  }

  /* 运行区监控 */
  .running-container {
    display: flex;
    flex-direction: column;
    gap: 14px;
  }

  .status-summary-card {
    background: #ffffff;
    border: 1px solid #edebe9;
    border-radius: 6px;
    padding: 14px;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
  }

  .status-main-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
  }

  .status-badge-title {
    display: flex;
    align-items: center;
    gap: 10px;
  }

  .status-pill {
    padding: 4px 10px;
    border-radius: 14px;
    font-size: 12px;
    font-weight: 600;
    display: inline-flex;
    align-items: center;
    gap: 6px;
  }

  .pill-running {
    background: #dbeeff;
    color: #005a9e;
  }
  .pill-completed {
    background: #dff6dd;
    color: #107c41;
  }
  .pill-error {
    background: #fed9cc;
    color: #a80000;
  }
  .pill-cancelled {
    background: #edebe9;
    color: #605e5c;
  }

  .job-id-text {
    font-size: 11px;
    color: #8a8886;
  }

  .status-actions {
    display: flex;
    gap: 8px;
  }

  .cancel-notice {
    font-size: 11px;
    color: #795b00;
    background: #fff4ce;
    padding: 4px 8px;
    border-radius: 4px;
  }

  .folder-notice {
    font-size: 11px;
    padding: 2px 4px;
  }

  /* 指标网格 */
  .metrics-grid {
    display: grid;
    grid-template-columns: repeat(7, 1fr);
    gap: 8px;
  }

  .metric-card {
    background: #f8f7f6;
    border: 1px solid #edebe9;
    border-radius: 4px;
    padding: 8px 6px;
    text-align: center;
    display: flex;
    flex-direction: column;
    gap: 2px;
  }

  .metric-num {
    font-size: 16px;
    font-weight: 700;
    color: #201f1e;
  }

  .metric-label {
    font-size: 10px;
    color: #605e5c;
  }

  .card-success .metric-num {
    color: #107c41;
  }
  .card-failed .metric-num {
    color: #a80000;
  }
  .card-blocked .metric-num {
    color: #c43e1c;
  }
  .card-cancelled .metric-num {
    color: #8a8886;
  }
  .card-pending .metric-num {
    color: #005a9e;
  }

  /* 逐文件处理表格 */
  .results-table-section {
    background: #ffffff;
    border: 1px solid #edebe9;
    border-radius: 6px;
    padding: 14px;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
  }

  .table-header-title {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 10px;
  }

  .table-header-title h3 {
    margin: 0;
    font-size: 13px;
    font-weight: 600;
    color: #201f1e;
  }

  .table-hint {
    font-size: 10px;
    color: #8a8886;
  }

  .results-list-box {
    max-height: 320px;
    overflow-y: auto;
    display: flex;
    flex-direction: column;
    gap: 6px;
  }

  .result-row-card {
    border: 1px solid #edebe9;
    border-radius: 4px;
    padding: 8px 10px;
    background: #faf9f8;
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  .status-success {
    border-left: 3px solid #107c41;
  }
  .status-failed {
    border-left: 3px solid #a80000;
    background: #fffcfb;
  }
  .status-blocked {
    border-left: 3px solid #d83b01;
    background: #fffdfb;
  }
  .status-running {
    border-left: 3px solid #0078d4;
    background: #fdfefe;
  }
  .status-cancelled {
    border-left: 3px solid #8a8886;
  }

  .row-main-line {
    display: flex;
    align-items: center;
    justify-content: space-between;
  }

  .row-file-badge {
    display: flex;
    align-items: center;
    gap: 6px;
    min-width: 0;
  }

  .file-num {
    font-size: 10px;
    color: #8a8886;
    font-weight: 600;
  }

  .result-file-name {
    font-weight: 600;
    color: #201f1e;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .status-tag {
    font-size: 11px;
    font-weight: 500;
    display: inline-flex;
    align-items: center;
    gap: 4px;
  }

  .tag-success {
    color: #107c41;
  }
  .tag-failed {
    color: #a80000;
  }
  .tag-blocked {
    color: #d83b01;
  }
  .tag-running {
    color: #0078d4;
  }
  .tag-cancelled {
    color: #605e5c;
  }
  .tag-pending {
    color: #8a8886;
  }

  .result-sub-line {
    display: flex;
    align-items: flex-start;
    gap: 6px;
    font-size: 10px;
    color: #605e5c;
  }

  .sub-label {
    flex-shrink: 0;
    font-weight: 500;
  }

  .sub-content {
    word-break: break-all;
  }

  .error-line {
    color: #a80000;
  }

  .tip-line {
    color: #8a8886;
    font-style: italic;
  }

  /* 按钮通用 */
  .btn-primary {
    background: #107c41;
    color: #ffffff;
    border: none;
    border-radius: 4px;
    padding: 6px 14px;
    font-size: 11px;
    font-weight: 500;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 6px;
  }
  .btn-primary:hover:not(:disabled) {
    background: #0b5a2f;
  }
  .btn-primary:disabled {
    background: #c8c6c4;
    cursor: not-allowed;
  }

  .btn-success {
    background: #107c41;
    color: #ffffff;
    border: none;
    border-radius: 4px;
    padding: 6px 16px;
    font-size: 12px;
    font-weight: 600;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 6px;
  }
  .btn-success:hover:not(:disabled) {
    background: #0b5a2f;
  }
  .btn-success:disabled {
    background: #c8c6c4;
    cursor: not-allowed;
  }

  .btn-secondary {
    background: #f3f2f1;
    color: #201f1e;
    border: 1px solid #d2d0ce;
    border-radius: 4px;
    padding: 6px 12px;
    font-size: 11px;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 4px;
  }
  .btn-secondary:hover {
    background: #edebe9;
  }

  .btn-danger-outline {
    background: #ffffff;
    color: #a80000;
    border: 1px solid #a80000;
    border-radius: 4px;
    padding: 6px 12px;
    font-size: 11px;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 4px;
  }
  .btn-danger-outline:hover:not(:disabled) {
    background: #fdf3f2;
  }

  .modal-footer {
    height: 44px;
    background: #fcfcfc;
    border-top: 1px solid #edebe9;
    padding: 0 16px;
    display: flex;
    align-items: center;
    justify-content: flex-end;
    flex-shrink: 0;
  }

  @keyframes spin {
    from {
      transform: rotate(0deg);
    }
    to {
      transform: rotate(360deg);
    }
  }

  .spin {
    animation: spin 1.5s linear infinite;
  }

  @media (max-width: 500px) {
    .form-grid-two {
      grid-template-columns: 1fr;
    }
    .metrics-grid {
      grid-template-columns: repeat(4, 1fr);
    }
  }

  @media (max-width: 360px) {
    .batch-modal-window {
      width: calc(100vw - 8px);
      height: calc(100vh - 12px);
    }
    .metrics-grid {
      grid-template-columns: repeat(2, 1fr);
    }
  }
</style>
