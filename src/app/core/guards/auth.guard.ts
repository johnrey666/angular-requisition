// src/app/core/guards/auth.guard.ts
import { Injectable, Inject, PLATFORM_ID } from '@angular/core';
import { CanActivate, Router, ActivatedRouteSnapshot, RouterStateSnapshot } from '@angular/router';
import { AuthService } from '../services/auth.service';
import { isPlatformBrowser } from '@angular/common';

@Injectable({
  providedIn: 'root'
})
export class AuthGuard implements CanActivate {
  constructor(
    private authService: AuthService,
    private router: Router,
    @Inject(PLATFORM_ID) private platformId: Object
  ) {}

  async canActivate(route: ActivatedRouteSnapshot, state: RouterStateSnapshot): Promise<boolean> {
    console.log('AuthGuard: Checking authentication for route:', state.url);
    
    // Wait for auth to be ready and get user
    const user = await this.authService.getCurrentUserPromise(8000);
    
    if (user) {
      console.log('AuthGuard: User authenticated, allowing access to:', state.url);
      return true;
    }

    console.log('AuthGuard: No user found, redirecting to landing from:', state.url);
    
    // Only store redirect URL in browser environment
    if (isPlatformBrowser(this.platformId)) {
      localStorage.setItem('redirectUrl', state.url);
    }
    
    await this.router.navigate(['/']);
    return false;
  }
}