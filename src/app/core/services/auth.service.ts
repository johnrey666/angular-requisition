// src/core/services/auth.service.ts
import { Injectable } from '@angular/core';
import { Auth, authState, signInWithEmailAndPassword, signOut, createUserWithEmailAndPassword } from '@angular/fire/auth';
import { BehaviorSubject, Observable } from 'rxjs';
import { map } from 'rxjs/operators';

@Injectable({ providedIn: 'root' })
export class AuthService {
  user$: BehaviorSubject<any> = new BehaviorSubject<any>(null);

  constructor(private auth: Auth) {
    // Subscribe to auth state changes
    authState(this.auth).subscribe((user) => {
      this.user$.next(user);
    });
  }

  // Get current user as observable
  getCurrentUser(): Observable<any> {
    return authState(this.auth);
  }

  // Get current user as promise
  getCurrentUserPromise(): Promise<any> {
    return Promise.resolve(this.auth.currentUser);
  }

  // Sign in with email and password
  signIn(email: string, password: string): Promise<any> {
    return signInWithEmailAndPassword(this.auth, email, password);
  }

  // Sign out
  signOut(): Promise<void> {
    return signOut(this.auth);
  }

  // Create new user (admin only)
  createUser(email: string, password: string): Promise<any> {
    return createUserWithEmailAndPassword(this.auth, email, password);
  }

  // Check if user is authenticated
  isAuthenticated(): Observable<boolean> {
    return authState(this.auth).pipe(
      map(user => !!user)
    );
  }

  // Get user ID
  getUserId(): string | null {
    return this.user$.value?.uid || null;
  }

  // Get user email
  getUserEmail(): string | null {
    return this.user$.value?.email || null;
  }
}