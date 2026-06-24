import { Component, inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ToastService, ToastType } from '../../services/toast.service';

@Component({
  selector: 'app-toast',
  standalone: true,
  imports: [CommonModule],
  template: `
    <div class="toast-container" aria-live="polite" aria-atomic="false">
      @for (toast of toastService.toasts(); track toast.id) {
        <div
          class="toast"
          [class.toast-success]="toast.type === 'success'"
          [class.toast-error]="toast.type === 'error'"
          [class.toast-info]="toast.type === 'info'"
          [class.toast-warning]="toast.type === 'warning'"
          role="alert"
        >
          <span class="toast-icon" aria-hidden="true">{{ getIcon(toast.type) }}</span>
          <span class="toast-message">{{ toast.message }}</span>
          <button
            type="button"
            class="toast-close"
            (click)="toastService.dismiss(toast.id)"
            aria-label="Dismiss notification"
          >
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <line x1="18" y1="6" x2="6" y2="18"/>
              <line x1="6" y1="6" x2="18" y2="18"/>
            </svg>
          </button>
        </div>
      }
    </div>
  `,
  styles: [`
    .toast-container {
      position: fixed;
      bottom: 1.25rem;
      right: 1.25rem;
      z-index: 10000;
      display: flex;
      flex-direction: column-reverse;
      gap: 0.625rem;
      max-width: min(420px, calc(100vw - 2rem));
      pointer-events: none;
    }

    .toast {
      display: flex;
      align-items: flex-start;
      gap: 0.75rem;
      padding: 0.875rem 1rem;
      border-radius: 0.75rem;
      background: hsl(var(--card));
      border: 1px solid hsl(var(--border));
      box-shadow:
        0 10px 40px rgb(0 0 0 / 0.12),
        0 2px 8px rgb(0 0 0 / 0.06);
      pointer-events: auto;
      animation: toast-in 0.28s cubic-bezier(0.16, 1, 0.3, 1);
      backdrop-filter: blur(12px);
    }

    .toast-icon {
      flex-shrink: 0;
      width: 1.375rem;
      height: 1.375rem;
      display: flex;
      align-items: center;
      justify-content: center;
      border-radius: 50%;
      font-size: 0.75rem;
      font-weight: 700;
    }

    .toast-message {
      flex: 1;
      font-size: 0.875rem;
      line-height: 1.45;
      color: hsl(var(--foreground));
      padding-top: 0.1rem;
    }

    .toast-close {
      flex-shrink: 0;
      display: flex;
      align-items: center;
      justify-content: center;
      width: 1.5rem;
      height: 1.5rem;
      border: none;
      border-radius: 0.375rem;
      background: transparent;
      color: hsl(var(--muted-foreground));
      cursor: pointer;
      transition: background 0.15s, color 0.15s;
      padding: 0;
    }

    .toast-close svg {
      width: 0.875rem;
      height: 0.875rem;
    }

    .toast-close:hover {
      background: hsl(var(--muted) / 0.6);
      color: hsl(var(--foreground));
    }

    .toast-success {
      border-color: hsl(142 76% 36% / 0.25);
    }

    .toast-success .toast-icon {
      background: hsl(142 76% 36% / 0.15);
      color: hsl(142 76% 32%);
    }

    .toast-error {
      border-color: hsl(0 84% 60% / 0.25);
    }

    .toast-error .toast-icon {
      background: hsl(0 84% 60% / 0.12);
      color: hsl(0 84% 45%);
    }

    .toast-info {
      border-color: hsl(217 91% 60% / 0.25);
    }

    .toast-info .toast-icon {
      background: hsl(217 91% 60% / 0.12);
      color: hsl(217 91% 50%);
    }

    .toast-warning {
      border-color: hsl(38 92% 50% / 0.3);
    }

    .toast-warning .toast-icon {
      background: hsl(38 92% 50% / 0.15);
      color: hsl(32 95% 40%);
    }

    @keyframes toast-in {
      from {
        opacity: 0;
        transform: translateY(0.75rem) scale(0.96);
      }
      to {
        opacity: 1;
        transform: translateY(0) scale(1);
      }
    }

    @media (max-width: 480px) {
      .toast-container {
        left: 1rem;
        right: 1rem;
        bottom: 1rem;
        max-width: none;
      }
    }
  `],
})
export class ToastComponent {
  readonly toastService = inject(ToastService);

  getIcon(type: ToastType): string {
    switch (type) {
      case 'success': return '✓';
      case 'error': return '✕';
      case 'warning': return '!';
      default: return 'i';
    }
  }
}
