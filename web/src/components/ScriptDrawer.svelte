<script lang="ts">
  import { X, Play, Trash2, Code2, ChevronDown, ChevronUp, FileCode } from 'lucide-svelte';
  import { bridge, type ScriptItem } from '../services/bridge';

  export let isOpen = false;
  export let onClose: () => void;
  export let onRunScript: (script: ScriptItem) => void;

  let scripts: ScriptItem[] = [];
  let isLoading = false;
  let expandedIndex: number | null = null;

  export async function refreshScripts() {
    isLoading = true;
    const res = await bridge.send<ScriptItem[]>('list_scripts');
    isLoading = false;
    if (res.ok && res.data) {
      scripts = res.data;
    }
  }

  $: if (isOpen) {
    refreshScripts();
  }

  async function handleDelete(fileName: string) {
    if (!confirm(`确认删除脚本 [${fileName}] 吗？`)) return;
    const res = await bridge.send('delete_script', { fileName });
    if (res.ok) {
      await refreshScripts();
    } else {
      alert('删除失败: ' + res.error);
    }
  }
</script>

{#if isOpen}
  <div class="drawer-backdrop" on:click|self={onClose} on:keydown={(e) => e.key === 'Escape' && onClose()} role="presentation">
    <div class="drawer-content" role="dialog" aria-modal="true" aria-label="我的脚本库">
      <div class="drawer-header">
        <div class="drawer-title">
          <Code2 size={16} color="#107C41" />
          <span>我的脚本库</span>
          <span class="script-count">({scripts.length})</span>
        </div>
        <button class="btn-icon" on:click={onClose}><X size={16} /></button>
      </div>

      <div class="drawer-subtitle">
        保存在本地目录: <code>%AppData%\LeeExcel\Scripts\</code>
      </div>

      <div class="drawer-body">
        {#if isLoading}
          <div class="empty-state">正在加载本地脚本...</div>
        {:else if scripts.length === 0}
          <div class="empty-state">
            <FileCode size={32} color="#c8c6c4" />
            <p>暂无已保存的脚本</p>
            <span class="empty-hint">在聊天窗口执行 VBA 后，点击“保存到我的脚本”即可在此持久化。</span>
          </div>
        {:else}
          <div class="script-list">
            {#each scripts as s, idx}
              <div class="script-item">
                <div class="item-header">
                  <div class="item-info">
                    <div class="item-name" title={s.name}>{s.name}</div>
                    <div class="item-meta">
                      <span>{s.createdAt}</span>
                      {#if s.description}
                        <span class="meta-dot">·</span>
                        <span class="item-desc" title={s.description}>{s.description}</span>
                      {/if}
                    </div>
                  </div>

                  <div class="item-actions">
                    <button
                      class="btn btn-sm btn-primary"
                      on:click={() => {
                        onClose();
                        onRunScript(s);
                      }}
                      title="立即对当前工作簿运行此脚本"
                    >
                      <Play size={12} />
                      <span>运行</span>
                    </button>
                    <button
                      class="btn-icon btn-sm-icon"
                      on:click={() => (expandedIndex = expandedIndex === idx ? null : idx)}
                      title="查看代码"
                    >
                      {#if expandedIndex === idx}
                        <ChevronUp size={14} />
                      {:else}
                        <ChevronDown size={14} />
                      {/if}
                    </button>
                    <button class="btn-icon btn-sm-icon" on:click={() => handleDelete(s.fileName)} title="删除">
                      <Trash2 size={14} color="#A80000" />
                    </button>
                  </div>
                </div>

                {#if expandedIndex === idx}
                  <div class="code-preview">
                    <pre><code>{s.code}</code></pre>
                  </div>
                {/if}
              </div>
            {/each}
          </div>
        {/if}
      </div>
    </div>
  </div>
{/if}

<style>
  .drawer-backdrop {
    position: fixed;
    top: 0;
    left: 0;
    width: 100vw;
    height: 100vh;
    background: rgba(0, 0, 0, 0.35);
    z-index: 90;
    display: flex;
    justify-content: flex-end;
  }

  .drawer-content {
    width: 360px;
    height: 100vh;
    background: white;
    box-shadow: -4px 0 16px rgba(0, 0, 0, 0.12);
    display: flex;
    flex-direction: column;
    animation: slideIn 0.2s ease-out;
  }

  @keyframes slideIn {
    from {
      transform: translateX(100%);
    }
    to {
      transform: translateX(0);
    }
  }

  .drawer-header {
    height: 48px;
    padding: 0 14px;
    border-bottom: 1px solid var(--office-border);
    display: flex;
    align-items: center;
    justify-content: space-between;
  }

  .drawer-title {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 13px;
    font-weight: 600;
  }

  .script-count {
    font-size: 12px;
    color: var(--office-muted);
    font-weight: normal;
  }

  .drawer-subtitle {
    padding: 6px 14px;
    font-size: 11px;
    color: var(--office-muted);
    background: #f8f8f8;
    border-bottom: 1px solid var(--office-border);
  }

  .drawer-subtitle code {
    font-family: Consolas, monospace;
    color: var(--office-text);
  }

  .drawer-body {
    flex: 1;
    overflow-y: auto;
    padding: 10px;
  }

  .empty-state {
    padding: 40px 16px;
    text-align: center;
    color: var(--office-muted);
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 8px;
  }

  .empty-hint {
    font-size: 11px;
    line-height: 1.4;
    max-width: 240px;
  }

  .script-list {
    display: flex;
    flex-direction: column;
    gap: 8px;
  }

  .script-item {
    border: 1px solid var(--office-border);
    border-radius: 4px;
    background: white;
    overflow: hidden;
    transition: border-color 0.15s;
  }

  .script-item:hover {
    border-color: #c8c6c4;
  }

  .item-header {
    padding: 10px 12px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
  }

  .item-info {
    min-width: 0;
    flex: 1;
  }

  .item-name {
    font-size: 12px;
    font-weight: 600;
    color: var(--office-text);
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .item-meta {
    font-size: 11px;
    color: var(--office-muted);
    display: flex;
    align-items: center;
    gap: 4px;
    margin-top: 2px;
  }

  .item-desc {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    max-width: 130px;
  }

  .meta-dot {
    color: #c8c6c4;
  }

  .item-actions {
    display: flex;
    align-items: center;
    gap: 4px;
    flex-shrink: 0;
  }

  .btn-sm-icon {
    padding: 4px;
  }

  .code-preview {
    padding: 8px 10px;
    background: #1e1e1e;
    color: #d4d4d4;
    border-top: 1px solid #333;
    max-height: 180px;
    overflow-y: auto;
  }

  .code-preview pre {
    font-family: Consolas, monospace;
    font-size: 11px;
    line-height: 1.4;
    white-space: pre-wrap;
  }
</style>
