import { writable } from 'svelte/store';

export interface ToastItem {
  id: string;
  type: 'success' | 'error' | 'info';
  message: string;
  duration?: number;
}

export const toastStore = writable<ToastItem[]>([]);

function addToast(type: 'success' | 'error' | 'info', message: string, duration = 3500) {
  const id = 'toast_' + Math.random().toString(36).slice(2, 9);
  const item: ToastItem = { id, type, message, duration };

  toastStore.update((items) => {
    const next = [...items, item];
    if (next.length > 3) {
      return next.slice(next.length - 3);
    }
    return next;
  });

  if (duration > 0) {
    setTimeout(() => {
      dismissToast(id);
    }, duration);
  }
}

export function dismissToast(id: string) {
  toastStore.update((items) => items.filter((i) => i.id !== id));
}

export const toast = {
  success: (msg: string) => addToast('success', msg),
  error: (msg: string) => addToast('error', msg),
  info: (msg: string) => addToast('info', msg),
  dismiss: dismissToast,
};
