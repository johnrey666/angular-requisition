// src/app/core/services/auth.service.ts
import { Injectable, Injector, runInInjectionContext, Inject, PLATFORM_ID } from '@angular/core';
import { Auth, authState, signInWithEmailAndPassword, signOut, createUserWithEmailAndPassword, User } from '@angular/fire/auth';
import { Observable, firstValueFrom, timeout, catchError, of, BehaviorSubject } from 'rxjs';
import { map, take, filter } from 'rxjs/operators';
import { isPlatformBrowser } from '@angular/common';

@Injectable({ providedIn: 'root' })
export class AuthService {
  private authState$: Observable<User | null>;
  private authReady = new BehaviorSubject<boolean>(false);
  private isBrowser: boolean;

  constructor(
    private auth: Auth, 
    private injector: Injector,
    @Inject(PLATFORM_ID) private platformId: Object
  ) {
    this.isBrowser = isPlatformBrowser(this.platformId);
    
    this.authState$ = new Observable(subscriber => {
      // Only subscribe to auth state in browser
      if (this.isBrowser) {
        runInInjectionContext(this.injector, () => {
          authState(this.auth).subscribe({
            next: (user) => {
              console.log('Auth state changed:', user?.email);
              subscriber.next(user);
              this.authReady.next(true);
            },
            error: (err) => {
              console.error('Auth state error:', err);
              subscriber.error(err);
              this.authReady.next(true);
            }
          });
        });
      } else {
        // On server, just emit null and mark as ready
        subscriber.next(null);
        this.authReady.next(true);
      }
    });
  }

  // Wait for auth to be ready
  async waitForAuth(timeoutMs: number = 3000): Promise<boolean> {
    // On server, just return true immediately
    if (!this.isBrowser) {
      return true;
    }
    
    if (this.auth.currentUser) {
      return true;
    }
    
    try {
      await firstValueFrom(
        this.authReady.pipe(
          filter(ready => ready === true),
          take(1),
          timeout(timeoutMs)
        )
      );
      return true;
    } catch {
      return false;
    }
  }

  // Get current user as observable
  getCurrentUserObservable(): Observable<User | null> {
    return this.authState$;
  }

  // Get current user as promise - improved with better error handling
  async getCurrentUserPromise(timeoutMs: number = 5000): Promise<User | null> {
    // On server, always return null
    if (!this.isBrowser) {
      return null;
    }
    
    // First check if we already have a user synchronously
    if (this.auth.currentUser) {
      console.log('Found current user synchronously:', this.auth.currentUser.email);
      return this.auth.currentUser;
    }
    
    // Wait for auth to be ready first
    await this.waitForAuth(3000);
    
    // Then get the user
    try {
      const user = await firstValueFrom(
        this.authState$.pipe(
          take(1),
          timeout(timeoutMs),
          catchError((error) => {
            console.log('Timeout or error waiting for auth state:', error);
            return of(null);
          })
        )
      );
      console.log('Auth state resolved:', user?.email);
      return user;
    } catch (err) {
      console.log('Error getting current user:', err);
      return null;
    }
  }

  // Get current user synchronously
  getCurrentUser() {
    if (!this.isBrowser) {
      return null;
    }
    return this.auth.currentUser;
  }

  // Sign in with email and password
  signIn(email: string, password: string): Promise<any> {
    console.log('Attempting to sign in:', email);
    return signInWithEmailAndPassword(this.auth, email, password);
  }

  // Sign out
  signOut(): Promise<void> {
    console.log('Signing out');
    if (this.isBrowser) {
      sessionStorage.removeItem('adminPassword');
    }
    this.authReady.next(false);
    return signOut(this.auth);
  }

  // Create new user (admin only)
  createUser(email: string, password: string): Promise<any> {
    return createUserWithEmailAndPassword(this.auth, email, password);
  }

  // Check if user is authenticated
  isAuthenticated(): Observable<boolean> {
    return this.authState$.pipe(
      map(user => !!user)
    );
  }

  // Get user ID
  getUserId(): string | null {
    if (!this.isBrowser) {
      return null;
    }
    return this.auth.currentUser?.uid || null;
  }

  // Get user email
  getUserEmail(): string | null {
    if (!this.isBrowser) {
      return null;
    }
    return this.auth.currentUser?.email || null;
  }
}