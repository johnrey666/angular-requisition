// src/app/dashboard/pages/users/users.component.ts
import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { AuthService } from '../../../core/services/auth.service';
import { UserService } from '../../../core/services/user.service';
import { Firestore, collection, query, getDocs, doc, deleteDoc, updateDoc } from '@angular/fire/firestore';
import { Timestamp } from 'firebase/firestore';
import { FormsModule } from '@angular/forms';

interface User {
  id: string;
  email: string;
  role: string;
  createdAt?: Timestamp | any;
}

@Component({
  selector: 'app-users',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
    <div class="page-content">
      <!-- Role Header -->
      <div class="role-header">
        <span class="role-badge role-admin">Admin Access</span>
        <span class="role-description">User Management Dashboard</span>
      </div>
      
      <!-- Notification -->
      <div class="snackbar" [class.show]="showNotification" [class]="'snackbar-' + notificationType">
        <span>{{ notificationMessage }}</span>
        <button class="snackbar-close" (click)="hideNotification()">✕</button>
      </div>
      
      <!-- Header -->
      <div class="page-header">
        <div class="header-title">
          <h1>User Management</h1>
          <span class="subtitle">Manage system users and permissions</span>
        </div>
        <div class="header-actions">
          <button class="btn btn-primary" (click)="openCreateUserModal()">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor">
              <path d="M12 5v14M5 12h14"/>
            </svg>
            Create New User
          </button>
        </div>
      </div>
      
      <!-- Users Table -->
      <div class="card" *ngIf="!isLoading">
        <div class="card-header">
          <h3>System Users</h3>
          <span class="user-count">{{ users.length }} total</span>
        </div>
        <div class="table-wrapper">
          <table>
            <thead>
              <tr>
                <th>Email</th>
                <th>Role</th>
                <th>Created</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              <tr *ngFor="let user of users">
                <td>{{ user.email }}</td>
                <td>
                  <select 
                    class="role-select" 
                    [value]="user.role" 
                    (change)="updateUserRole(user, $event)"
                    [disabled]="user.email === 'admin@gmail.com'">
                    <option value="user">User</option>
                    <option value="store">Store</option>
                    <option value="production">Production</option>
                    <option value="procurement">Procurement</option>
                    <option value="admin">Admin</option>
                  </select>
                </td>
                <td>{{ formatDate(user.createdAt) }}</td>
                <td>
                  <button 
                    class="btn-icon delete" 
                    title="Delete User" 
                    (click)="deleteUser(user)" 
                    *ngIf="user.email !== 'admin@gmail.com'">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                      <polyline points="3 6 5 6 21 6"/>
                      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/>
                      <line x1="10" y1="11" x2="10" y2="17"/>
                      <line x1="14" y1="11" x2="14" y2="17"/>
                    </svg>
                  </button>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- Loading State -->
      <div *ngIf="isLoading" class="loading-state">
        <div class="spinner"></div>
        <p>Loading users...</p>
      </div>

      <!-- Create User Modal -->
      <div class="modal" [class.active]="showCreateModal">
        <div class="modal-backdrop" (click)="closeCreateUserModal()"></div>
        <div class="modal-container">
          <div class="modal-header">
            <h3>Create New User</h3>
            <button class="close" (click)="closeCreateUserModal()">✕</button>
          </div>
          
          <div class="modal-body">
            <form #createUserForm="ngForm">
              <div class="form-group">
                <label>Email *</label>
                <input 
                  type="email" 
                  [(ngModel)]="newUser.email" 
                  name="email" 
                  required 
                  email
                  #emailField="ngModel"
                  placeholder="user@example.com">
                <span class="error" *ngIf="emailField.invalid && emailField.touched">
                  Please enter a valid email
                </span>
              </div>

              <div class="form-group">
                <label>Password *</label>
                <input 
                  type="password" 
                  [(ngModel)]="newUser.password" 
                  name="password" 
                  required 
                  minlength="6"
                  #passwordField="ngModel"
                  placeholder="••••••••">
                <span class="error" *ngIf="passwordField.invalid && passwordField.touched">
                  Password must be at least 6 characters
                </span>
              </div>

              <div class="form-group">
                <label>Role *</label>
                <select [(ngModel)]="newUser.role" name="role" required>
                  <option value="user">User</option>
                  <option value="store">Store</option>
                  <option value="production">Production</option>
                  <option value="procurement">Procurement</option>
                  <option value="admin">Admin</option>
                </select>
              </div>

              <div class="form-group" *ngIf="createError || createSuccess">
                <div class="alert error" *ngIf="createError">{{ createError }}</div>
                <div class="alert success" *ngIf="createSuccess">{{ createSuccess }}</div>
              </div>
            </form>
          </div>

          <div class="modal-footer">
            <button class="btn btn-secondary" (click)="closeCreateUserModal()">Cancel</button>
            <button 
              class="btn btn-primary" 
              (click)="createUser()" 
              [disabled]="isCreating || !createUserForm.form.valid">
              <span *ngIf="isCreating" class="spinner-small"></span>
              {{ isCreating ? 'Creating...' : 'Create User' }}
            </button>
          </div>
        </div>
      </div>
    </div>
  `,
  styles: [`
    .page-content {
      padding: 24px;
    }
    .role-header {
      margin-bottom: 24px;
      padding: 12px 16px;
      background: var(--surface-color);
      border-radius: 8px;
      display: flex;
      align-items: center;
      gap: 12px;
    }
    .role-badge {
      padding: 4px 12px;
      border-radius: 16px;
      font-size: 12px;
      font-weight: 600;
      text-transform: uppercase;
    }
    .role-admin {
      background: #ef4444;
      color: white;
    }
    .role-description {
      color: var(--text-secondary);
    }
    .page-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 24px;
    }
    .header-title h1 {
      margin: 0;
      font-size: 24px;
      font-weight: 600;
    }
    .subtitle {
      color: var(--text-secondary);
      font-size: 14px;
    }
    .btn {
      padding: 8px 16px;
      border-radius: 6px;
      border: none;
      font-size: 14px;
      font-weight: 500;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      background: var(--surface-color);
      border: 1px solid var(--border-color);
    }
    .btn-primary {
      background: #3b82f6;
      color: white;
      border: none;
    }
    .btn-primary:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }
    .btn-secondary {
      background: var(--surface-color);
      border: 1px solid var(--border-color);
    }
    .card {
      background: var(--surface-color);
      border-radius: 8px;
      overflow: hidden;
      border: 1px solid var(--border-color);
    }
    .card-header {
      padding: 16px;
      border-bottom: 1px solid var(--border-color);
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .card-header h3 {
      margin: 0;
    }
    .user-count {
      color: var(--text-secondary);
      font-size: 14px;
    }
    .table-wrapper {
      overflow-x: auto;
    }
    table {
      width: 100%;
      border-collapse: collapse;
    }
    th {
      text-align: left;
      padding: 12px 16px;
      background: var(--background-color);
      font-weight: 500;
      color: var(--text-secondary);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    td {
      padding: 12px 16px;
      border-bottom: 1px solid var(--border-color);
    }
    .role-select {
      padding: 4px 8px;
      border-radius: 4px;
      border: 1px solid var(--border-color);
      background: var(--surface-color);
      font-size: 13px;
    }
    .btn-icon {
      padding: 6px;
      border: none;
      background: none;
      cursor: pointer;
      opacity: 0.6;
      margin: 0 4px;
    }
    .btn-icon:hover {
      opacity: 1;
    }
    .btn-icon.delete:hover {
      color: #ef4444;
    }
    .loading-state {
      text-align: center;
      padding: 48px;
      color: var(--text-secondary);
    }
    .spinner {
      width: 40px;
      height: 40px;
      border: 3px solid var(--border-color);
      border-top-color: #3b82f6;
      border-radius: 50%;
      animation: spin 1s linear infinite;
      margin: 0 auto 16px;
    }
    .spinner-small {
      width: 16px;
      height: 16px;
      border: 2px solid rgba(255,255,255,0.3);
      border-top-color: white;
      border-radius: 50%;
      animation: spin 1s linear infinite;
      display: inline-block;
      margin-right: 8px;
    }
    @keyframes spin {
      to { transform: rotate(360deg); }
    }

    /* Modal Styles */
    .modal {
      display: none;
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      z-index: 1000;
    }
    .modal.active {
      display: block;
    }
    .modal-backdrop {
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(0, 0, 0, 0.5);
    }
    .modal-container {
      position: absolute;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      background: var(--surface-color);
      border-radius: 8px;
      width: 90%;
      max-width: 500px;
      max-height: 90vh;
      overflow-y: auto;
      box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1);
    }
    .modal-header {
      padding: 16px 20px;
      border-bottom: 1px solid var(--border-color);
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .modal-header h3 {
      margin: 0;
    }
    .close {
      background: none;
      border: none;
      font-size: 20px;
      cursor: pointer;
      color: var(--text-secondary);
    }
    .modal-body {
      padding: 20px;
    }
    .modal-footer {
      padding: 16px 20px;
      border-top: 1px solid var(--border-color);
      display: flex;
      justify-content: flex-end;
      gap: 12px;
    }

    /* Form Styles */
    .form-group {
      margin-bottom: 16px;
    }
    label {
      display: block;
      margin-bottom: 6px;
      font-weight: 500;
      font-size: 14px;
    }
    input, select {
      width: 100%;
      padding: 10px;
      border: 1px solid var(--border-color);
      border-radius: 6px;
      background: var(--surface-color);
      font-size: 14px;
    }
    input:focus, select:focus {
      outline: none;
      border-color: #3b82f6;
    }
    .error {
      color: #ef4444;
      font-size: 12px;
      margin-top: 4px;
      display: block;
    }
    .alert {
      padding: 12px;
      border-radius: 6px;
      font-size: 14px;
    }
    .alert.error {
      background: #fee2e2;
      color: #991b1b;
      border: 1px solid #fecaca;
    }
    .alert.success {
      background: #dcfce7;
      color: #166534;
      border: 1px solid #bbf7d0;
    }

    /* Snackbar */
    .snackbar {
      position: fixed;
      bottom: 24px;
      right: 24px;
      background: var(--surface-color);
      border: 1px solid var(--border-color);
      border-radius: 8px;
      padding: 12px 16px;
      display: flex;
      align-items: center;
      gap: 12px;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
      transform: translateY(100px);
      opacity: 0;
      transition: all 0.3s ease;
      z-index: 1100;
    }
    .snackbar.show {
      transform: translateY(0);
      opacity: 1;
    }
    .snackbar.snackbar-success {
      background: #22c55e;
      color: white;
      border-color: #16a34a;
    }
    .snackbar.snackbar-error {
      background: #ef4444;
      color: white;
      border-color: #dc2626;
    }
    .snackbar.snackbar-info {
      background: #3b82f6;
      color: white;
      border-color: #2563eb;
    }
    .snackbar-close {
      background: none;
      border: none;
      color: currentColor;
      cursor: pointer;
      opacity: 0.7;
    }
    .snackbar-close:hover {
      opacity: 1;
    }
  `]
})
export class UsersComponent implements OnInit {
  users: User[] = [];
  isLoading = true;
  
  // Create user modal
  showCreateModal = false;
  isCreating = false;
  createError: string | null = null;
  createSuccess: string | null = null;
  
  newUser = {
    email: '',
    password: '',
    role: 'user'
  };

  // Notification
  showNotification = false;
  notificationMessage = '';
  notificationType: 'success' | 'error' | 'info' = 'info';
  notificationTimeout: any;

  constructor(
    private authService: AuthService,
    private userService: UserService,
    private firestore: Firestore
  ) {}

  async ngOnInit() {
    await this.loadUsers();
  }

  async loadUsers() {
    try {
      const usersRef = collection(this.firestore, 'users');
      const snapshot = await getDocs(usersRef);
      
      this.users = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      })) as User[];
      
      // Sort users by email
      this.users.sort((a, b) => a.email.localeCompare(b.email));
      
      console.log('Loaded users:', this.users);
    } catch (err) {
      console.error('Failed to load users:', err);
      this.showMessage('Failed to load users', 'error');
    } finally {
      this.isLoading = false;
    }
  }

  formatDate(date: any): string {
    if (!date) return 'N/A';
    
    // Handle Firestore Timestamp
    if (date && typeof date.toDate === 'function') {
      return date.toDate().toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
      });
    }
    
    // Handle Date object
    if (date instanceof Date) {
      return date.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
      });
    }
    
    // Handle string date
    if (typeof date === 'string') {
      return new Date(date).toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
      });
    }
    
    return 'N/A';
  }

  openCreateUserModal() {
    this.showCreateModal = true;
    this.createError = null;
    this.createSuccess = null;
    this.newUser = {
      email: '',
      password: '',
      role: 'user'
    };
  }

  closeCreateUserModal() {
    this.showCreateModal = false;
  }

  async createUser() {
    if (!this.newUser.email || !this.newUser.password || !this.newUser.role) {
      this.createError = 'Please fill all required fields';
      return;
    }

    this.isCreating = true;
    this.createError = null;
    this.createSuccess = null;

    try {
      console.log('Creating user:', this.newUser.email);
      
      await this.userService.createUserAccount(
        this.newUser.email,
        this.newUser.password,
        this.newUser.role
      );
      
      this.createSuccess = 'User created successfully!';
      
      // Reload users list
      await this.loadUsers();
      
      // Show success message
      this.showMessage(`User ${this.newUser.email} created successfully`, 'success');
      
      // Close modal after delay
      setTimeout(() => {
        this.closeCreateUserModal();
      }, 1500);
      
    } catch (err: any) {
      console.error('Create user failed:', err);
      
      if (err.code === 'auth/email-already-in-use') {
        this.createError = 'This email is already registered.';
      } else if (err.code === 'auth/weak-password') {
        this.createError = 'Password must be at least 6 characters.';
      } else {
        this.createError = err?.message || 'Failed to create user. Please try again.';
      }
    } finally {
      this.isCreating = false;
    }
  }

  async updateUserRole(user: User, event: any) {
    const newRole = event.target.value;
    
    if (user.role === newRole) return;
    
    if (!confirm(`Change role for ${user.email} from ${user.role} to ${newRole}?`)) {
      // Reset select to previous value
      event.target.value = user.role;
      return;
    }

    try {
      const userRef = doc(this.firestore, 'users', user.id);
      await updateDoc(userRef, {
        role: newRole,
        updatedAt: new Date()
      });
      
      user.role = newRole;
      this.showMessage(`Role updated for ${user.email}`, 'success');
    } catch (err) {
      console.error('Failed to update role:', err);
      this.showMessage('Failed to update role', 'error');
      // Reset select to previous value
      event.target.value = user.role;
    }
  }

  async deleteUser(user: User) {
    if (user.email === 'admin@gmail.com') {
      this.showMessage('Cannot delete the main admin account', 'error');
      return;
    }

    if (!confirm(`Are you sure you want to delete user ${user.email}? This action cannot be undone.`)) {
      return;
    }

    try {
      const userRef = doc(this.firestore, 'users', user.id);
      await deleteDoc(userRef);
      
      // Note: Deleting from Firestore doesn't delete from Auth
      // You'll need a Cloud Function to also delete from Auth
      
      this.users = this.users.filter(u => u.id !== user.id);
      this.showMessage(`User ${user.email} deleted`, 'success');
    } catch (err) {
      console.error('Failed to delete user:', err);
      this.showMessage('Failed to delete user', 'error');
    }
  }

  showMessage(message: string, type: 'success' | 'error' | 'info' = 'info') {
    if (this.notificationTimeout) {
      clearTimeout(this.notificationTimeout);
    }
    
    this.notificationMessage = message;
    this.notificationType = type;
    this.showNotification = true;
    
    this.notificationTimeout = setTimeout(() => {
      this.showNotification = false;
    }, 3000);
  }

  hideNotification() {
    this.showNotification = false;
    if (this.notificationTimeout) {
      clearTimeout(this.notificationTimeout);
    }
  }
}