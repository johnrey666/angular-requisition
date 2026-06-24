import { Component, Inject, PLATFORM_ID } from '@angular/core';
import { Router, RouterOutlet, NavigationEnd } from '@angular/router';
import { filter, take } from 'rxjs/operators';
import { isPlatformBrowser } from '@angular/common';
import { ThemeService } from './core/services/theme.service';
import { LoaderComponent } from './core/components/loader/loader.component';
import { ToastComponent } from './core/components/toast/toast.component';
import { LoaderService } from './core/services/loader.service';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet, LoaderComponent, ToastComponent],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css'
})
export class AppComponent {
  constructor(
    private themeService: ThemeService,
    private loader: LoaderService,
    private router: Router,
    @Inject(PLATFORM_ID) private platformId: Object
  ) {
    if (isPlatformBrowser(this.platformId)) {
      this.loader.setBootstrapping(true);
      this.router.events.pipe(
        filter((e): e is NavigationEnd => e instanceof NavigationEnd),
        take(1)
      ).subscribe(() => this.loader.setBootstrapping(false));
    }
  }
}