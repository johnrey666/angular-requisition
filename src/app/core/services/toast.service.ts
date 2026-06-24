import { Injectable, signal } from '@angular/core';

export type ToastType = 'success' | 'error' | 'info' | 'warning';

export interface Toast {
  id: number;
  message: string;
  type: ToastType;
}

@Injectable({ providedIn: 'root' })
export class ToastService {
  private nextId = 0;
  private readonly timeouts = new Map<number, ReturnType<typeof setTimeout>>();
  readonly toasts = signal<Toast[]>([]);

  show(message: string, type: ToastType = 'info', duration = 4000): void {
    const id = ++this.nextId;
    this.toasts.update((items) => [...items, { id, message, type }]);

    if (duration > 0) {
      const timeout = setTimeout(() => this.dismiss(id), duration);
      this.timeouts.set(id, timeout);
    }
  }

  success(message: string, duration = 4000): void {
    this.show(message, 'success', duration);
  }

  error(message: string, duration = 5000): void {
    this.show(message, 'error', duration);
  }

  info(message: string, duration = 4000): void {
    this.show(message, 'info', duration);
  }

  warning(message: string, duration = 4500): void {
    this.show(message, 'warning', duration);
  }

  dismiss(id: number): void {
    const timeout = this.timeouts.get(id);
    if (timeout) {
      clearTimeout(timeout);
      this.timeouts.delete(id);
    }
    this.toasts.update((items) => items.filter((t) => t.id !== id));
  }

  dismissAll(): void {
    this.timeouts.forEach((timeout) => clearTimeout(timeout));
    this.timeouts.clear();
    this.toasts.set([]);
  }
}
