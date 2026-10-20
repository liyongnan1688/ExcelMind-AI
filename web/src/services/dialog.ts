import { writable } from 'svelte/store';

export interface ConfirmOptions {
  title: string;
  message: string;
  confirmText?: string;
  cancelText?: string;
  danger?: boolean;
}

export interface DialogState extends ConfirmOptions {
  isOpen: boolean;
  resolve?: (value: boolean) => void;
}

export const dialogStore = writable<DialogState>({
  isOpen: false,
  title: '',
  message: '',
  confirmText: '确定',
  cancelText: '取消',
  danger: false,
});

export function showConfirm(options: ConfirmOptions): Promise<boolean> {
  return new Promise<boolean>((resolve) => {
    dialogStore.set({
      isOpen: true,
      title: options.title,
      message: options.message,
      confirmText: options.confirmText || '确定',
      cancelText: options.cancelText || '取消',
      danger: !!options.danger,
      resolve,
    });
  });
}

export function closeConfirm(result: boolean) {
  dialogStore.update((state) => {
    if (state.resolve) {
      state.resolve(result);
    }
    return {
      ...state,
      isOpen: false,
      resolve: undefined,
    };
  });
}
