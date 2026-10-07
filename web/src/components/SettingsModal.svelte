<script lang="ts">
  import { X, Check, KeyRound, Globe, Cpu, ShieldCheck, Download, AlertCircle, RefreshCw, FileText, CheckCircle2 } from 'lucide-svelte';
  import { PRESET_PROVIDERS, loadLlmConfig, saveLlmConfig, type LlmConfig } from '../services/config';
  import { bridge, type DiagnosticsPreviewResult, type DiagnosticsExportResult } from '../services/bridge';

  export let onClose: () => void;

  let activeTab: 'api' | 'diagnostics' = 'api';
  let config: LlmConfig = loadLlmConfig();
  let savedAlert = false;

  // 诊断状态
  let diagLoading = false;
  let diagPreview: DiagnosticsPreviewResult | null = null;
  let diagExporting = false;
  let diagExportResult: DiagnosticsExportResult | null = null;
  let diagError: string | null = null;

  function handleSelectPreset(presetId: string) {
    const p = PRESET_PROVIDERS.find((x) => x.id === presetId);
    if (p) {
      config.provider = p.id;
      config.baseUrl = p.baseUrl;
      config.model = p.model;
    }
  }

  function handleSave() {
    saveLlmConfig(config);
    savedAlert = true;
    setTimeout(() => {
      savedAlert = false;
      onClose();
    }, 800);
  }

  async function loadDiagnostics() {
    diagLoading = true;
    diagError = null;
    diagExportResult = null;
    try {
      diagPreview = await bridge.previewDiagnostics();
      if (!diagPreview.ok) {
        diagError = diagPreview.error || '获取诊断预览失败';
      }
    } catch (e: any) {
      diagError = e?.message || '获取诊断信息异常';
    } finally {
      diagLoading = false;
    }
  }

  function switchTab(tab: 'api' | 'diagnostics') {
    activeTab = tab;
    if (tab === 'diagnostics' && !diagPreview) {
      loadDiagnostics();
    }
  }

  async function handleExportDiagnostics() {
    diagExporting = true;
    diagError = null;
    try {
      const res = await bridge.exportDiagnostics();
      if (res.ok) {
        diagExportResult = res;
      } else {
        if (res.error && !res.error.includes('取消')) {
          diagError = res.error;
        }
      }
    } catch (e: any) {
      diagError = e?.message || '导出诊断包遇到异常';
    } finally {
      diagExporting = false;
    }
  }
</script>

<div class="modal-backdrop" on:click|self={onClose} on:keydown={(e) => e.key === 'Escape' && onClose()} role="presentation">
  <div class="modal-content {activeTab === 'diagnostics' ? 'wide' : ''}" role="dialog" aria-modal="true" aria-labelledby="modal-settings-title">
    <div class="modal-header">
      <div class="header-left">
        <div id="modal-settings-title" class="modal-title">设置与环境诊断</div>
        <div class="tab-pills">
          <button class="tab-pill {activeTab === 'api' ? 'active' : ''}" on:click={() => switchTab('api')}>
            大模型 API
          </button>
          <button class="tab-pill {activeTab === 'diagnostics' ? 'active' : ''}" on:click={() => switchTab('diagnostics')}>
            脱敏诊断与导出
          </button>
        </div>
      </div>
      <button class="btn-icon" on:click={onClose} aria-label="关闭"><X size={16} /></button>
    </div>

    {#if activeTab === 'api'}
      <div class="modal-body">
        <div class="form-group">
          <span class="form-label" id="lbl-preset">选择厂商预设</span>
          <div class="preset-grid" aria-labelledby="lbl-preset">
            {#each PRESET_PROVIDERS as p}
              <button
                class="preset-btn {config.provider === p.id ? 'active' : ''}"
                on:click={() => handleSelectPreset(p.id)}
                type="button"
              >
                {p.name}
              </button>
            {/each}
          </div>
        </div>

        <div class="form-group">
          <label for="cfg-base-url"><Globe size={13} style="margin-right: 4px; vertical-align: -2px;" /> API 接口地址 (Base URL)</label>
          <input id="cfg-base-url" type="text" bind:value={config.baseUrl} placeholder="https://api.deepseek.com/v1" />
          <span class="hint">输入兼容 OpenAI 的根路径（结尾无需 /chat/completions）</span>
        </div>

        <div class="form-group">
          <label for="cfg-api-key"><KeyRound size={13} style="margin-right: 4px; vertical-align: -2px;" /> API Key 密钥</label>
          <input id="cfg-api-key" type="password" bind:value={config.apiKey} placeholder="sk-..." />
          <span class="hint">密钥安全存储在本地浏览器 LocalStorage 中，诊断导出绝对不打包</span>
        </div>

        <div class="form-group">
          <label for="cfg-model"><Cpu size={13} style="margin-right: 4px; vertical-align: -2px;" /> 模型名称 (Model Name)</label>
          <input id="cfg-model" type="text" bind:value={config.model} placeholder="deepseek-chat" />
        </div>

        <div class="form-group">
          <label for="cfg-temperature">生成温度 (Temperature，可选)</label>
          <input
            id="cfg-temperature"
            type="number"
            step="0.1"
            min="0"
            max="2"
            bind:value={config.temperature}
            placeholder="留空表示使用模型服务默认值"
          />
          <span class="hint">留空时不强制传参；若填写建议 0.3 ~ 0.7 之间</span>
        </div>

        <div class="form-group">
          <label for="cfg-max-tokens">最大输出 Token 预算 (Max Tokens)</label>
          <input
            id="cfg-max-tokens"
            type="number"
            step="1024"
            min="512"
            max="65536"
            bind:value={config.maxTokens}
            placeholder="留空表示不限制（采用服务端默认限制）"
          />
          <span class="hint">留空时不发送 max_tokens 字段；用户填写则作为客户端输出上限发送</span>
        </div>

        <div class="form-group">
          <label for="cfg-thinking-mode">深度思考 / 推理模式 (Thinking Mode)</label>
          <select id="cfg-thinking-mode" bind:value={config.thinkingMode} class="select-input">
            <option value="auto">自动 (不发送 thinking 字段，完全由模型服务端决定)</option>
            <option value="disabled">关闭思考链 (发送 type: disabled，防止思考耗尽预算)</option>
            <option value="budget">指定思考预算 (发送 type: enabled，可选配置 budget_tokens)</option>
          </select>
        </div>

        {#if config.thinkingMode === 'budget'}
          <div class="form-group">
            <label for="cfg-thinking-budget">思考链预算 Token 数 (Thinking Budget)</label>
            <input
              id="cfg-thinking-budget"
              type="number"
              step="512"
              min="512"
              max="16384"
              bind:value={config.thinkingBudget}
              placeholder="留空表示不限制（采用服务端默认）"
            />
            <span class="hint">留空时不发送 budget_tokens；用户填写时才发送具体预算数值</span>
          </div>
        {/if}
      </div>

      <div class="modal-footer">
        {#if savedAlert}
          <span class="save-tip"><Check size={14} /> 保存成功！</span>
        {:else}
          <span></span>
        {/if}
        <div class="actions">
          <button class="btn" on:click={onClose}>取消</button>
          <button class="btn btn-primary" on:click={handleSave}>保存配置</button>
        </div>
      </div>
    {:else}
      <!-- 诊断与环境导出选项卡 -->
      <div class="modal-body diag-body">
        {#if diagLoading}
          <div class="loading-state">
            <RefreshCw size={18} class="spin" />
            <span>正在收集系统白名单与诊断数据...</span>
          </div>
        {:else}
          {#if diagError}
            <div class="alert-box error">
              <AlertCircle size={14} />
              <span>{diagError}</span>
            </div>
          {/if}

          {#if diagExportResult}
            <div class="alert-box success">
              <CheckCircle2 size={16} />
              <div class="export-success-details">
                <strong>脱敏诊断包导出成功！</strong>
                <div class="file-path">{diagExportResult.zipFilePath}</div>
                <div class="file-meta">
                  大小: {Math.round(diagExportResult.zipSizeBytes / 1024)} KB | 
                  SHA-256: <code>{diagExportResult.sha256?.substring(0, 16)}...</code> | 
                  日志行数: {diagExportResult.totalLogLines}
                </div>
              </div>
            </div>
          {/if}

          {#if diagPreview && diagPreview.summary}
            <!-- 环境信息概览卡 -->
            <div class="diag-card">
              <div class="card-title">系统与宿主环境概览 (只读白名单)</div>
              <div class="env-grid">
                <div class="env-item"><span class="k">插件版本:</span> <span class="v">{diagPreview.summary.appVersion}</span></div>
                <div class="env-item"><span class="k">操作系统:</span> <span class="v">{diagPreview.summary.osVersion} ({diagPreview.summary.osArchitecture})</span></div>
                <div class="env-item"><span class="k">Excel 宿主:</span> <span class="v">{diagPreview.summary.excelVersion} ({diagPreview.summary.excelBitness})</span></div>
                <div class="env-item"><span class="k">WebView2:</span> <span class="v">{diagPreview.summary.webView2Version}</span></div>
                <div class="env-item"><span class="k">活动工作簿:</span> <span class="v">{diagPreview.summary.hasActiveWorkbook ? diagPreview.summary.activeWorkbookMaskedName : '无'}</span></div>
                <div class="env-item"><span class="k">最后故障阶段:</span> <span class="v">{diagPreview.summary.lastFailureStage}</span></div>
              </div>
            </div>

            <!-- 白名单收集范围 -->
            <div class="diag-card">
              <div class="card-title"><CheckCircle2 size={14} color="#107C41" /> 诊断包收集范围 (仅含以下脱敏文件)</div>
              <ul class="clean-list">
                <li><code>diagnostics_summary.json</code>：系统环境、CLR、Excel 及 WebView2 运行时只读白名单</li>
                <li><code>diagnostics.log</code>：已脱敏的运行时日志（最近 {diagPreview.sanitizedLogLinesCount} 行，敏感行自动剔除）</li>
                <li><code>manifest.json</code>：包元数据清单与生成时间（版本 1.0）</li>
              </ul>
            </div>

            <!-- 严格排除范围 -->
            <div class="diag-card excluded">
              <div class="card-title"><ShieldCheck size={14} color="#5c2d91" /> 严格排除范围 (100% 物理隔离绝不导出)</div>
              <div class="excluded-tags">
                <span class="tag">大模型 API Key 密钥</span>
                <span class="tag">DPAPI 本地凭据密文</span>
                <span class="tag">用户会话与提示词正文</span>
                <span class="tag">宏库源码 .bas 正文</span>
                <span class="tag">宏运行实际参数与载荷</span>
                <span class="tag">工作簿及单元格业务数据</span>
                <span class="tag">历史物理快照文件</span>
                <span class="tag">外部 HTTP 响应缓存</span>
              </div>
            </div>

            <!-- 日志脱敏预览 -->
            {#if diagPreview.sanitizedLogPreview && diagPreview.sanitizedLogPreview.length > 0}
              <div class="diag-card">
                <div class="card-title"><FileText size={14} /> 脱敏日志预览 (前 {diagPreview.sanitizedLogPreview.length} 行)</div>
                <pre class="log-preview">{diagPreview.sanitizedLogPreview.join('\n')}</pre>
                {#if diagPreview.omittedSensitiveLinesCount > 0}
                  <span class="hint-warn">注意：已自动剔除 {diagPreview.omittedSensitiveLinesCount} 行潜在高危敏感日志。</span>
                {/if}
              </div>
            {/if}
          {/if}
        {/if}
      </div>

      <div class="modal-footer">
        <span class="safe-promise"><ShieldCheck size={13} /> 诊断包纯本地打包，绝不自动联网上传</span>
        <div class="actions">
          <button class="btn" on:click={loadDiagnostics} disabled={diagLoading || diagExporting}>
            <RefreshCw size={13} class={diagLoading ? 'spin' : ''} /> 刷新
          </button>
          <button class="btn btn-primary" on:click={handleExportDiagnostics} disabled={diagLoading || diagExporting}>
            <Download size={13} /> {diagExporting ? '正在打包...' : '导出脱敏诊断包 (.zip)'}
          </button>
        </div>
      </div>
    {/if}
  </div>
</div>

<style>
  .modal-backdrop {
    position: fixed;
    top: 0;
    left: 0;
    width: 100vw;
    height: 100vh;
    background: rgba(0, 0, 0, 0.42);
    backdrop-filter: blur(2px);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 100;
    padding: 14px;
    box-sizing: border-box;
  }

  @media (max-width: 360px) {
    .modal-backdrop {
      padding: 6px;
    }
  }

  .modal-content {
    width: min(440px, 100%);
    max-height: calc(100vh - 24px);
    background: #ffffff;
    border-radius: var(--office-radius-lg);
    box-shadow: var(--office-shadow-lg);
    display: flex;
    flex-direction: column;
    overflow: hidden;
    border: 1px solid var(--office-border);
    animation: modalPop 0.15s ease-out;
    box-sizing: border-box;
  }

  .modal-content.wide {
    width: min(520px, 100%);
  }

  @keyframes modalPop {
    from {
      opacity: 0;
      transform: scale(0.96);
    }
    to {
      opacity: 1;
      transform: scale(1);
    }
  }

  .modal-header {
    height: 52px;
    padding: 0 16px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 1px solid var(--office-border);
    background: #faf9f8;
  }

  .header-left {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  .modal-title {
    font-size: 13px;
    font-weight: 600;
    color: var(--office-text);
  }

  .tab-pills {
    display: flex;
    gap: 4px;
  }

  .tab-pill {
    background: transparent;
    border: none;
    font-size: 11px;
    padding: 2px 8px;
    border-radius: 4px;
    cursor: pointer;
    color: var(--office-text-secondary);
  }

  .tab-pill.active {
    background: #ffffff;
    color: var(--excel-green);
    font-weight: 600;
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.08);
  }

  .modal-body {
    padding: 16px;
    overflow-y: auto;
    display: flex;
    flex-direction: column;
    gap: 14px;
    font-size: 12px;
  }

  .diag-body {
    gap: 10px;
    background: #faf9f8;
  }

  .diag-card {
    background: #ffffff;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-sm);
    padding: 10px 12px;
    display: flex;
    flex-direction: column;
    gap: 6px;
  }

  .diag-card.excluded {
    background: #fdfafc;
    border-color: #ebd7f7;
  }

  .card-title {
    font-size: 11px;
    font-weight: 600;
    color: var(--office-text);
    display: flex;
    align-items: center;
    gap: 6px;
  }

  .env-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 4px 10px;
    font-size: 11px;
  }

  .env-item .k {
    color: var(--office-text-secondary);
  }

  .env-item .v {
    font-weight: 500;
    color: var(--office-text);
  }

  .clean-list {
    margin: 0;
    padding-left: 18px;
    font-size: 11px;
    color: var(--office-text);
    line-height: 1.5;
  }

  .clean-list code {
    background: #f3f2f1;
    padding: 1px 4px;
    border-radius: 2px;
    font-size: 10px;
  }

  .excluded-tags {
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
  }

  .tag {
    background: #f3e9f9;
    color: #5c2d91;
    font-size: 10px;
    padding: 2px 6px;
    border-radius: 3px;
    border: 1px solid #ebd7f7;
  }

  .log-preview {
    margin: 0;
    background: #201f1e;
    color: #f3f2f1;
    font-size: 10px;
    padding: 6px 8px;
    border-radius: 3px;
    max-height: 120px;
    overflow-y: auto;
    font-family: var(--font-family-mono, monospace);
    white-space: pre-wrap;
    word-break: break-all;
  }

  .hint-warn {
    font-size: 10px;
    color: #a80000;
  }

  .alert-box {
    display: flex;
    align-items: flex-start;
    gap: 8px;
    padding: 8px 12px;
    border-radius: var(--office-radius-sm);
    font-size: 11px;
  }

  .alert-box.success {
    background: #e8f5e9;
    border: 1px solid #c8e6c9;
    color: #1b5e20;
  }

  .alert-box.error {
    background: #fde8e8;
    border: 1px solid #f8b4b4;
    color: #9b1c1c;
  }

  .export-success-details {
    display: flex;
    flex-direction: column;
    gap: 2px;
  }

  .file-path {
    font-family: var(--font-family-mono, monospace);
    font-size: 10px;
    word-break: break-all;
    background: rgba(255, 255, 255, 0.6);
    padding: 2px 4px;
    border-radius: 2px;
  }

  .file-meta {
    font-size: 10px;
    color: #2e7d32;
  }

  .file-meta code {
    background: rgba(255, 255, 255, 0.6);
    padding: 1px 3px;
    border-radius: 2px;
  }

  .loading-state {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    padding: 30px;
    color: var(--office-text-secondary);
  }

  .form-group {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }

  .form-label,
  label {
    font-size: 12px;
    font-weight: 500;
    color: var(--office-text);
  }

  input,
  .select-input {
    height: 30px;
    padding: 0 8px;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-sm);
    font-size: 12px;
    background: #ffffff;
    color: var(--office-text);
  }

  input:focus,
  .select-input:focus {
    border-color: var(--excel-green);
    outline: none;
  }

  .hint {
    font-size: 11px;
    color: var(--office-text-secondary);
  }

  .preset-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 6px;
  }

  .preset-btn {
    height: 28px;
    background: #f3f2f1;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-sm);
    font-size: 11px;
    cursor: pointer;
    color: var(--office-text);
    transition: all 0.1s;
  }

  .preset-btn:hover {
    background: #edebe9;
  }

  .preset-btn.active {
    background: var(--excel-green-light);
    border-color: var(--excel-green);
    color: var(--excel-green);
    font-weight: 600;
  }

  .modal-footer {
    height: 48px;
    padding: 0 16px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-top: 1px solid var(--office-border);
    background: #faf9f8;
  }

  .save-tip {
    color: var(--excel-green);
    font-size: 12px;
    display: flex;
    align-items: center;
    gap: 4px;
  }

  .safe-promise {
    font-size: 11px;
    color: #605e5c;
    display: flex;
    align-items: center;
    gap: 4px;
  }

  .actions {
    display: flex;
    gap: 8px;
  }

  .btn {
    height: 28px;
    padding: 0 12px;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-sm);
    background: #ffffff;
    font-size: 12px;
    cursor: pointer;
    color: var(--office-text);
    display: inline-flex;
    align-items: center;
    gap: 4px;
  }

  .btn:hover:not(:disabled) {
    background: #f3f2f1;
  }

  .btn:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }

  .btn-primary {
    background: var(--excel-green);
    border-color: var(--excel-green);
    color: #ffffff;
    font-weight: 500;
  }

  .btn-primary:hover:not(:disabled) {
    background: var(--excel-green-dark, #0b5a30);
    border-color: var(--excel-green-dark, #0b5a30);
  }

  .btn-icon {
    width: 28px;
    height: 28px;
    border: none;
    background: transparent;
    border-radius: var(--office-radius-sm);
    display: flex;
    align-items: center;
    justify-content: center;
    cursor: pointer;
    color: var(--office-text);
  }

  .btn-icon:hover {
    background: #edebe9;
  }

  :global(.spin) {
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
