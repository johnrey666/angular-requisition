import { Component, inject, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { Router, NavigationStart, NavigationEnd } from '@angular/router';
import { LoaderService } from '../../services/loader.service';

@Component({
  selector: 'app-loader',
  standalone: true,
  imports: [CommonModule],
  template: `
    @if (loader.visible()) {
      <div class="loader-overlay" [class.content-only]="loader.contentOnly() && !loader.bootstrapping()" role="status" aria-live="polite" aria-label="Loading">
        <div class="loader-spinner"></div>
        @if (loader.message() && !loader.bootstrapping()) {
          <span class="loader-text">{{ loader.message() }}</span>
        }
      </div>
    }
  `,
  styles: [`
    .loader-overlay {
      position: fixed;
      inset: 0;
      z-index: 9999;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 0.75rem;
      background: hsl(var(--background));
      pointer-events: auto;
    }

    .loader-overlay.content-only {
      left: 16rem;
      top: 4rem;
      right: 0;
      bottom: 0;
    }

    :host-context(body.sidebar-collapsed) .loader-overlay.content-only {
      left: 4.5rem;
    }

    .loader-spinner {
      width: 18px;
      height: 18px;
      border: 2px solid hsl(var(--muted));
      border-top-color: hsl(var(--primary));
      border-radius: 50%;
      animation: spin 0.5s linear infinite;
    }

    .loader-text {
      font-size: 0.75rem;
      color: hsl(var(--muted-foreground));
    }

    @keyframes spin {
      to { transform: rotate(360deg); }
    }
  `]
})
export class LoaderComponent implements OnInit, OnDestroy {
  loader = inject(LoaderService);
  private router = inject(Router);
  private bootstrapDone = false;
  private navSub: any;

  ngOnInit() {
    this.navSub = this.router.events.subscribe(e => {
      if (e instanceof NavigationStart) {
        if (this.bootstrapDone) {
          this.loader.show();
        }
      } else if (e instanceof NavigationEnd) {
        if (!this.bootstrapDone) {
          this.bootstrapDone = true;
          this.loader.setBootstrapping(false);
        }
        this.loader.setContentOnly(e.urlAfterRedirects.startsWith('/dashboard'));
        if (this.bootstrapDone) {
          this.loader.hide();
        }
      }
    });
  }

  ngOnDestroy() {
    this.navSub?.unsubscribe();
  }
}
