<script lang="ts">
  import { dialogStore, closeConfirm } from '../services/dialog';
  import { AlertTriangle } from 'lucide-svelte';

  let state = $state({
    isOpen: false,
    title: '',
    message: '',
    confirmText: '确定',
    cancelText: '取消',
    danger: false,
  });

  let dialogElement: HTMLDivElement | null = $state(null);
  let confirmButtonElement: HTMLButtonElement | null = $state(null);

  dialogStore.subscribe((val) => {
    state = val;
    if (val.isOpen) {
      setTimeout(() => {
        if (confirmButtonElement) {
          confirmButtonElement.focus();
        }
      }, 30);
    }
  });

  function handleKeyDown(e: KeyboardEvent) {
    if (!state.isOpen) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      closeConfirm(false);
    } else if (e.key === 'Enter') {
      e.preventDefault();
      closeConfirm(true);
    } else if (e.key === 'Tab' && dialogElement) {
      // Focus lock / trap
      const focusables = dialogElement.querySelectorAll<HTMLElement>(
        'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'
      );
      if (focusables.length > 0) {
        const first = focusables[0];
        const last = focusables[focusables.length - 1];
        if (e.shiftKey && document.activeElement === first) {
          e.preventDefault();
          last.focus();
        } else if (!e.shiftKey && document.activeElement === last) {
          e.preventDefault();
          first.focus();
        }
      }
    }
  }
</script>

<svelte:window onkeydown={handleKeyDown} />

{#if state.isOpen}
  <div
    class="dialog-backdrop"
    role="presentation"
    onclick={() => closeConfirm(false)}
  >
    <div
      class="dialog-card"
      role="alertdialog"
      aria-modal="true"
      aria-labelledby="confirm-dialog-title"
      aria-describedby="confirm-dialog-message"
      tabindex="-1"
      bind:this={dialogElement}
      onclick={(e) => e.stopPropagation()}
      onkeydown={(e) => e.stopPropagation()}
    >
      <div class="dialog-header">
        {#if state.danger}
          <div class="dialog-icon-danger">
            <AlertTriangle size={16} />
          </div>
        {/if}
        <h3 id="confirm-dialog-title" class="dialog-title">{state.title}</h3>
      </div>

      <div id="confirm-dialog-message" class="dialog-message">
        {state.message}
      </div>

      <div class="dialog-actions">
        <button
          type="button"
          class="dialog-btn dialog-btn-cancel"
          onclick={() => closeConfirm(false)}
        >
          {state.cancelText || '取消'}
        </button>
        <button
          type="button"
          class="dialog-btn {state.danger ? 'dialog-btn-danger' : 'dialog-btn-primary'}"
          bind:this={confirmButtonElement}
          onclick={() => closeConfirm(true)}
        >
          {state.confirmText || '确定'}
        </button>
      </div>
    </div>
  </div>
{/if}

<style>
  .dialog-backdrop {
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background-color: rgba(0, 0, 0, 0.4);
    z-index: 10000;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 16px;
  }

  .dialog-card {
    background-color: var(--office-card);
    border: 1px solid var(--office-border);
    border-radius: var(--office-radius);
    box-shadow: var(--office-shadow-lg);
    width: 100%;
    max-width: 360px;
    padding: 16px;
    display: flex;
    flex-direction: column;
    gap: 12px;
    animation: dialog-pop 0.15s ease-out;
  }

  .dialog-header {
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .dialog-icon-danger {
    color: var(--office-danger);
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
  }

  .dialog-title {
    font-size: var(--font-size-base);
    font-weight: 600;
    color: var(--office-text);
    margin: 0;
    line-height: var(--line-height-tight);
  }

  .dialog-message {
    font-size: var(--font-size-sm);
    color: var(--office-text-secondary);
    line-height: var(--line-height-normal);
    white-space: pre-line;
    word-break: break-word;
  }

  .dialog-actions {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 4px;
  }

  .dialog-btn {
    padding: 5px 14px;
    font-size: var(--font-size-sm);
    font-weight: 500;
    border-radius: var(--office-radius-sm);
    cursor: pointer;
    border: 1px solid transparent;
    transition: background-color 0.1s, border-color 0.1s;
  }

  .dialog-btn-cancel {
    background-color: var(--office-card);
    border-color: var(--office-border-strong);
    color: var(--office-text);
  }

  .dialog-btn-cancel:hover {
    background-color: var(--office-hover);
  }

  .dialog-btn-primary {
    background-color: var(--excel-green);
    color: #ffffff;
  }

  .dialog-btn-primary:hover {
    background-color: var(--excel-dark);
  }

  .dialog-btn-danger {
    background-color: var(--office-danger);
    color: #ffffff;
  }

  .dialog-btn-danger:hover {
    background-color: var(--office-danger-dark);
  }

  @keyframes dialog-pop {
    from {
      opacity: 0;
      transform: scale(0.96);
    }
    to {
      opacity: 1;
      transform: scale(1);
    }
  }
</style>
