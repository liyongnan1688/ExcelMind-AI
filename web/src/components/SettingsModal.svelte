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

<div class="modal-backdrop" on:click|self={onClose}>
  <div class="modal-content">
    <div class="modal-header">
      <div class="modal-title">大模型 API 配置</div>
      <button class="btn-icon" on:click={onClose}><X size={16} /></button>
    </div>

    <div class="modal-body">
      <div class="form-group">
        <label>选择厂商预设</label>
        <div class="preset-grid">
          {#each PRESET_PROVIDERS as p}
            <button
              class="preset-btn {config.provider === p.id ? 'active' : ''}"
              on:click={() => handleSelectPreset(p.id)}
            >
              {p.name}
            </button>
          {/each}
        </div>
      </div>

      <div class="form-group">
        <label><Globe size={13} style="margin-right: 4px; vertical-align: -2px;" /> API 接口地址 (Base URL)</label>
        <input type="text" bind:value={config.baseUrl} placeholder="https://api.deepseek.com/v1" />
        <span class="hint">输入兼容 OpenAI 的根路径（结尾无需 /chat/completions）</span>
      </div>

      <div class="form-group">
        <label><KeyRound size={13} style="margin-right: 4px; vertical-align: -2px;" /> API Key 密钥</label>
        <input type="password" bind:value={config.apiKey} placeholder="sk-..." />
        <span class="hint">密钥安全存储在本地浏览器 LocalStorage 中</span>
      </div>

      <div class="form-group">
        <label><Cpu size={13} style="margin-right: 4px; vertical-align: -2px;" /> 模型名称 (Model Name)</label>
        <input type="text" bind:value={config.model} placeholder="deepseek-chat" />
      </div>

      <div class="form-group">
        <label>生成温度 (Temperature，可选)</label>
        <input
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
        <label>最大输出 Token 预算 (Max Tokens)</label>
        <input
          type="number"
          step="1024"
          min="2048"
          max="65536"
          bind:value={config.maxTokens}
          placeholder="默认 16384 (充足预算避免复杂多步骤看板中途被截断)"
        />
        <span class="hint">为复杂大表与多过程 VBA 保留充足生成空间，默认 16384</span>
      </div>

      <div class="form-group">
        <label>深度思考 / 推理模式 (Thinking Mode)</label>
        <select bind:value={config.thinkingMode} class="select-input">
          <option value="auto">自动 (完全由模型端默认决定)</option>
          <option value="disabled">关闭思考链 (Token 全部分配给代码正文，防止思考耗尽预算)</option>
          <option value="budget">限制思考预算 (指定思考 Token 上限，兼顾规划与输出)</option>
        </select>
      </div>

      {#if config.thinkingMode === 'budget'}
        <div class="form-group">
          <label>思考链预算 Token 数 (Thinking Budget)</label>
          <input
            type="number"
            step="512"
            min="512"
            max="16384"
            bind:value={config.thinkingBudget}
            placeholder="默认 2048"
          />
          <span class="hint">限制思考链占用的最大 Token 数，剩余预算全部留给 VBA 代码正文</span>
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
    background: rgba(0, 0, 0, 0.4);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 100;
  }

  .modal-content {
    width: 380px;
    background: white;
    border-radius: 6px;
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.16);
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }

  .modal-header {
    height: 44px;
    padding: 0 16px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 1px solid var(--office-border);
  }

  .modal-title {
    font-size: 14px;
    font-weight: 600;
    color: var(--office-text);
  }

  .modal-body {
    padding: 16px;
    display: flex;
    flex-direction: column;
    gap: 14px;
  }

  .form-group {
    display: flex;
    flex-direction: column;
    gap: 4px;
  }

  label {
    font-size: 12px;
    font-weight: 600;
    color: var(--office-text);
  }

  input, .select-input {
    height: 32px;
    padding: 0 10px;
    font-size: 12px;
    border: 1px solid var(--office-border);
    border-radius: 4px;
    outline: none;
    transition: border-color 0.15s;
    background: #fff;
  }

  input:focus {
    border-color: var(--excel-green);
    box-shadow: 0 0 0 1px var(--excel-green);
  }

  .hint {
    font-size: 11px;
    color: var(--office-muted);
  }

  .preset-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 6px;
  }

  .preset-btn {
    padding: 6px 8px;
    font-size: 11px;
    text-align: left;
    background: #f8f8f8;
    border: 1px solid var(--office-border);
    border-radius: 4px;
    cursor: pointer;
    transition: all 0.15s;
  }

  .preset-btn:hover {
    background: #edebe9;
  }

  .preset-btn.active {
    background: var(--excel-light);
    border-color: var(--excel-green);
    color: var(--excel-green);
    font-weight: 600;
  }

  .modal-footer {
    padding: 12px 16px;
    background: #faf9f8;
    border-top: 1px solid var(--office-border);
    display: flex;
    align-items: center;
    justify-content: space-between;
  }

  .save-tip {
    font-size: 12px;
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
