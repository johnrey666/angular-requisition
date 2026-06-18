import { Injectable, signal, computed } from '@angular/core';

@Injectable({ providedIn: 'root' })
export class LoaderService {
  private count = 0;
  readonly loading = signal(false);
  readonly message = signal<string | null>(null);
  readonly bootstrapping = signal(false);
  readonly contentOnly = signal(false);

  readonly visible = computed(() => this.loading() || this.bootstrapping());

  show(message?: string): void {
    this.count++;
    this.loading.set(true);
    this.message.set(message ?? null);
  }

  hide(): void {
    this.count = Math.max(0, this.count - 1);
    if (this.count === 0) {
      this.loading.set(false);
      this.message.set(null);
    }
  }

  hideAll(): void {
    this.count = 0;
    this.loading.set(false);
    this.message.set(null);
  }

  setBootstrapping(value: boolean): void {
    this.bootstrapping.set(value);
    if (!value && typeof document !== 'undefined') {
      const el = document.getElementById('app-loader');
      if (el) el.remove();
    }
  }

  setContentOnly(value: boolean): void {
    this.contentOnly.set(value);
  }
}
