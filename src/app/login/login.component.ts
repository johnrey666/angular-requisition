import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormBuilder, FormGroup, ReactiveFormsModule, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { AuthService } from '../core/services/auth.service';
import { UserService } from '../core/services/user.service';
import { Firestore, doc, getDoc, setDoc } from '@angular/fire/firestore';

@Component({
  selector: 'app-login',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule],
  templateUrl: './login.component.html',
  styleUrl: './login.component.css',
})
export class LoginComponent {
  loginForm: FormGroup;
  isSubmitting = false;
  showPassword = false;
  authError: string | null = null;

  constructor(
    private fb: FormBuilder,
    private router: Router,
    private authService: AuthService,
    private userService: UserService,
    private firestore: Firestore
  ) {
    this.loginForm = this.fb.nonNullable.group({
      email: ['', [Validators.required, Validators.email]],
      password: ['', [Validators.required, Validators.minLength(6)]],
      rememberMe: [false],
    });
  }

  get email() { return this.loginForm.get('email'); }
  get password() { return this.loginForm.get('password'); }
  togglePassword(): void { this.showPassword = !this.showPassword; }

  async onSubmit(): Promise<void> {
    if (this.loginForm.invalid) {
      this.loginForm.markAllAsTouched();
      return;
    }

    this.isSubmitting = true;
    this.authError = null;

    const { email, password } = this.loginForm.getRawValue();

    try {
      const res = await this.authService.signIn(email, password);
      const user = res.user;

      if (user) {
        // Ensure user document exists with a role before proceeding
        const role = await this.ensureUserDocument(user);
        console.log('User role after ensureUserDocument:', role);

        if (role === 'admin') {
          this.userService.storeAdminPassword(password);
        }

        this.isSubmitting = false;

        switch (role) {
          case 'store':
            await this.router.navigate(['/dashboard/store']);
            break;
          case 'production':
            await this.router.navigate(['/dashboard/production']);
            break;
          case 'procurement':
            await this.router.navigate(['/dashboard/procurement']);
            break;
          case 'admin':
            await this.router.navigate(['/dashboard/users']);
            break;
          default:
            await this.router.navigate(['/dashboard']);
        }
      }
    } catch (err: any) {
      console.error('Sign in error', err);
      this.isSubmitting = false;
      this.mapAuthError(err);
    }
  }

  /**
   * Ensures the user document exists in Firestore with a valid role.
   * If the document doesn't exist, creates it with the default 'user' role.
   * Returns the user's role.
   */
  private async ensureUserDocument(user: any): Promise<string> {
    const userRef = doc(this.firestore, 'users', user.uid);
    
    try {
      const userDoc = await getDoc(userRef);

      if (userDoc.exists()) {
        const data = userDoc.data() as any;
        const role = data['role'];
        
        // Verify role field actually exists and is valid
        if (role && typeof role === 'string' && role.trim() !== '') {
          console.log('User doc found with role:', role);
          return role;
        }
        
        // Doc exists but role is missing — patch it
        console.warn('User doc exists but has no role, patching with default...');
        await setDoc(userRef, { 
          role: 'user',
          email: user.email,
          updated_at: new Date().toISOString()
        }, { merge: true });
        return 'user';
      }

      // Document doesn't exist at all — create it
      console.warn('User doc does not exist, creating with default role...');
      await setDoc(userRef, {
        email: user.email,
        role: 'user',
        created_at: new Date().toISOString(),
        updated_at: new Date().toISOString()
      });
      return 'user';

    } catch (err) {
      console.error('ensureUserDocument failed:', err);
      // Even if this fails, don't block login — return a safe default
      return 'user';
    }
  }

  private mapAuthError(err: any): void {
    const code = err?.code || '';
    if (code === 'auth/user-not-found' || code === 'auth/invalid-credential') {
      this.authError = 'No account found with this email';
    } else if (code === 'auth/wrong-password') {
      this.authError = 'Incorrect password';
    } else if (code === 'auth/invalid-email') {
      this.authError = 'Invalid email format';
    } else if (code === 'auth/too-many-requests') {
      this.authError = 'Too many failed attempts. Please try again later';
    } else {
      this.authError = err?.message || 'Sign in failed';
    }
  }
}