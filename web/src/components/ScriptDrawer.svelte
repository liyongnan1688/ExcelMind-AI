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
        <button class="btn-icon" on:click={onClose} aria-label="关闭"><X size={16} /></button>
      </div>

      <div class="drawer-subtitle">
        保存在本地目录: <code>%AppData%\ExcelMindAI\Scripts\</code>
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
                      title={expandedIndex === idx ? '收起源码' : '查看源码'}
                    >
                      {#if expandedIndex === idx}
                        <ChevronUp size={14} />
                      {:else}
                        <ChevronDown size={14} />
                      {/if}
                    </button>
                    <button class="btn-icon btn-sm-icon" on:click={() => handleDelete(s.fileName)} title="删除脚本">
                      <Trash2 size={14} color="#C42B1C" />
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
    background: rgba(0, 0, 0, 0.38);
    backdrop-filter: blur(2px);
    z-index: 90;
    display: flex;
    justify-content: flex-end;
  }

  .drawer-content {
    width: 100%;
    height: 100vh;
    background: #ffffff;
    box-shadow: -4px 0 20px rgba(0, 0, 0, 0.15);
    display: flex;
    flex-direction: column;
    animation: slideIn 0.2s cubic-bezier(0.16, 1, 0.3, 1);
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
    flex-shrink: 0;
  }

  .drawer-title {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: var(--font-size-md);
    font-weight: 600;
    color: var(--office-text);
  }

  .script-count {
    font-size: var(--font-size-xs);
    color: var(--office-muted);
    font-weight: normal;
  }

  .drawer-subtitle {
    padding: 6px 14px;
    font-size: var(--font-size-xs);
    color: var(--office-muted);
    background: var(--office-card-subtle);
    border-bottom: 1px solid var(--office-border);
    flex-shrink: 0;
  }

  .drawer-subtitle code {
    font-family: var(--font-family-code);
    color: var(--office-text);
    background: #edebe9;
    padding: 1px 4px;
    border-radius: var(--office-radius-xs);
  }

  .drawer-body {
    flex: 1;
    overflow-y: auto;
    padding: 12px;
  }

  .empty-state {
    padding: 40px 16px;
    text-align: center;
    color: var(--office-muted);
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 8px;
    font-size: var(--font-size-sm);
  }

  .empty-hint {
    font-size: var(--font-size-xs);
    line-height: var(--line-height-normal);
    max-width: 240px;
    color: var(--office-dim);
  }

  .script-list {
    display: flex;
    flex-direction: column;
    gap: 8px;
  }

  .script-item {
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    background: #ffffff;
    overflow: hidden;
    transition: all 0.15s ease;
  }

  .script-item:hover {
    border-color: var(--office-border-strong);
    box-shadow: var(--office-shadow-sm);
  }

  .item-header {
    padding: 9px 12px;
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
    font-size: var(--font-size-sm);
    font-weight: 600;
    color: var(--office-text);
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    line-height: 1.3;
  }

  .item-meta {
    font-size: var(--font-size-xs);
    color: var(--office-muted);
    display: flex;
    align-items: center;
    gap: 4px;
    margin-top: 2px;
    line-height: 1.3;
  }

  .item-desc {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    max-width: 140px;
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
    border-radius: var(--office-radius-xs);
  }

  .code-preview {
    padding: 8px 12px;
    background: #1e1e1e;
    color: #d4d4d4;
    border-top: 1px solid #333333;
    max-height: 200px;
    overflow-y: auto;
  }

  .code-preview pre {
    font-family: var(--font-family-code);
    font-size: var(--font-size-xs);
    line-height: var(--line-height-normal);
    white-space: pre-wrap;
    word-break: break-all;
  }
</style>
