<script lang="ts">
  import { toastStore, dismissToast } from '../services/toast';
  import { CheckCircle2, XCircle, Info, X } from 'lucide-svelte';

  let items = $state([]);
  toastStore.subscribe((val) => {
    items = val;
  });
</script>

{#if items.length > 0}
  <div class="toast-container" role="status" aria-live="polite">
    {#each items as item (item.id)}
      <div class="toast-card toast-{item.type}">
        <div class="toast-icon">
          {#if item.type === 'success'}
            <CheckCircle2 size={14} class="text-excel" />
          {:else if item.type === 'error'}
            <XCircle size={14} class="text-danger" />
          {:else}
            <Info size={14} class="text-info" />
          {/if}
        </div>
        <div class="toast-message">{item.message}</div>
        <button
          type="button"
          class="toast-close"
          onclick={() => dismissToast(item.id)}
          aria-label="关闭提示"
        >
          <X size={12} />
        </button>
      </div>
    {/each}
  </div>
{/if}

<style>
  .toast-container {
    position: fixed;
    bottom: 16px;
    left: 50%;
    transform: translateX(-50%);
    z-index: 9999;
    display: flex;
    flex-direction: column;
    gap: 8px;
    width: max-content;
    max-width: calc(100vw - 32px);
    pointer-events: none;
  }

  .toast-card {
    pointer-events: auto;
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 8px 12px;
    background-color: var(--office-card);
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius-sm);
    box-shadow: var(--office-shadow-lg);
    font-size: var(--font-size-sm);
    color: var(--office-text);
    animation: toast-fade-in 0.15s ease-out;
  }

  .toast-card.toast-success {
    border-color: var(--excel-light-border);
    background-color: var(--office-card);
  }

  .toast-card.toast-error {
    border-color: var(--office-danger-border);
    background-color: var(--office-danger-light);
  }

  .toast-card.toast-info {
    border-color: var(--office-blue-border);
    background-color: var(--office-blue-light);
  }

  .toast-icon {
    display: flex;
    align-items: center;
    flex-shrink: 0;
  }

  :global(.toast-icon .text-excel) {
    color: var(--excel-green);
  }

  :global(.toast-icon .text-danger) {
    color: var(--office-danger);
  }

  :global(.toast-icon .text-info) {
    color: var(--office-blue);
  }

  .toast-message {
    line-height: var(--line-height-normal);
    word-break: break-word;
    font-size: var(--font-size-sm);
  }

  .toast-close {
    border: none;
    background: transparent;
    color: var(--office-muted);
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 2px;
    border-radius: var(--office-radius-xs);
    margin-left: 4px;
    flex-shrink: 0;
  }

  .toast-close:hover {
    color: var(--office-text);
    background-color: var(--office-hover);
  }

  @keyframes toast-fade-in {
    from {
      opacity: 0;
      transform: translateY(6px);
    }
    to {
      opacity: 1;
      transform: translateY(0);
    }
  }
</style>
