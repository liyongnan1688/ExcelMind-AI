<script lang="ts">
  import { X, Check, KeyRound, Globe, Cpu } from 'lucide-svelte';
  import { PRESET_PROVIDERS, loadLlmConfig, saveLlmConfig, type LlmConfig } from '../services/config';

  export let onClose: () => void;

  let config: LlmConfig = loadLlmConfig();
  let savedAlert = false;

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
</script>

<div class="modal-backdrop" on:click|self={onClose} on:keydown={(e) => e.key === 'Escape' && onClose()} role="presentation">
  <div class="modal-content" role="dialog" aria-modal="true" aria-labelledby="modal-settings-title">
    <div class="modal-header">
      <div id="modal-settings-title" class="modal-title">大模型 API 配置</div>
      <button class="btn-icon" on:click={onClose} aria-label="关闭"><X size={16} /></button>
    </div>

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
        <span class="hint">密钥安全存储在本地浏览器 LocalStorage 中</span>
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
  }

  .modal-content {
    width: min(420px, calc(100vw - 28px));
    max-height: calc(100vh - 36px);
    background: #ffffff;
    border-radius: var(--office-radius-lg);
    box-shadow: var(--office-shadow-lg);
    display: flex;
    flex-direction: column;
    overflow: hidden;
    border: 1px solid var(--office-border);
    animation: modalPop 0.15s ease-out;
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
    height: 46px;
    padding: 0 16px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 1px solid var(--office-border);
    flex-shrink: 0;
    background: #ffffff;
  }

  .modal-title {
    font-size: var(--font-size-md);
    font-weight: 600;
    color: var(--office-text);
  }

  .modal-body {
    padding: 16px;
    display: flex;
    flex-direction: column;
    gap: 13px;
    overflow-y: auto;
    flex: 1;
  }

  .form-group {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  label, .form-label {
    font-size: var(--font-size-sm);
    font-weight: 600;
    color: var(--office-text);
  }

  input, .select-input {
    height: 32px;
    padding: 0 10px;
    font-size: var(--font-size-sm);
    font-family: inherit;
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-sm);
    outline: none;
    transition: all 0.15s ease;
    background: #ffffff;
    color: var(--office-text);
  }

  input:focus, .select-input:focus {
    border-color: var(--excel-green);
    box-shadow: 0 0 0 2px rgba(16, 124, 65, 0.15);
  }

  .hint {
    font-size: var(--font-size-xs);
    color: var(--office-muted);
    line-height: 1.35;
  }

  .preset-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 6px;
  }

  .preset-btn {
    padding: 6px 10px;
    font-size: var(--font-size-xs);
    font-family: inherit;
    text-align: left;
    background: var(--office-card-subtle);
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-sm);
    cursor: pointer;
    transition: all 0.15s ease;
    color: var(--office-text-secondary);
  }

  .preset-btn:hover {
    background: var(--office-hover);
    border-color: var(--office-border-strong);
  }

  .preset-btn.active {
    background: var(--excel-light);
    border-color: var(--excel-green);
    color: var(--excel-green);
    font-weight: 600;
  }

  .modal-footer {
    padding: 10px 16px;
    background: var(--office-card-subtle);
    border-top: 1px solid var(--office-border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-shrink: 0;
  }

  .save-tip {
    font-size: var(--font-size-sm);
    color: var(--excel-green);
    display: flex;
    align-items: center;
    gap: 4px;
    font-weight: 500;
  }

  .actions {
    display: flex;
    gap: 8px;
  }
</style>
