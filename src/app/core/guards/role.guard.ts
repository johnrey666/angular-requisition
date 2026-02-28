// src/app/core/guards/role.guard.ts
import { Injectable, Inject, PLATFORM_ID } from '@angular/core';
import { CanActivate, ActivatedRouteSnapshot, Router, RouterStateSnapshot } from '@angular/router';
import { AuthService } from '../services/auth.service';
import { Firestore, doc, getDoc } from '@angular/fire/firestore';
import { isPlatformBrowser } from '@angular/common';

@Injectable({
  providedIn: 'root'
})
export class RoleGuard implements CanActivate {
  constructor(
    private authService: AuthService,
    private firestore: Firestore,
    private router: Router,
    @Inject(PLATFORM_ID) private platformId: Object
  ) {}

  async canActivate(route: ActivatedRouteSnapshot, state: RouterStateSnapshot): Promise<boolean> {
    console.log('RoleGuard: Checking role for route:', state.url);
    
    // Wait for auth to be ready and get user
    const user = await this.authService.getCurrentUserPromise(8000);
    
    if (!user) {
      console.log('RoleGuard: No user found, redirecting to landing');
      await this.router.navigate(['/']);
      return false;
    }

    try {
      console.log('RoleGuard: Checking role for user:', user.uid);
      const userDocRef = doc(this.firestore, 'users', user.uid);
      const userDoc = await getDoc(userDocRef);
      
      let userRole = 'user';
      
      if (userDoc.exists()) {
        const data = userDoc.data();
        userRole = data['role'] || 'user';
        console.log('RoleGuard: User role found:', userRole);
      } else {
        console.log('RoleGuard: No user document found, using default role');
      }

      const allowedRoles = route.data['roles'] as Array<string>;
      
      if (allowedRoles && allowedRoles.includes(userRole)) {
        console.log('RoleGuard: Access granted for role:', userRole);
        return true;
      }

      console.log('RoleGuard: User role', userRole, 'not allowed for route:', state.url);
      await this.router.navigate(['/dashboard']);
      return false;
    } catch (err) {
      console.error('RoleGuard: Error checking role:', err);
      await this.router.navigate(['/']);
      return false;
    }
  }
}