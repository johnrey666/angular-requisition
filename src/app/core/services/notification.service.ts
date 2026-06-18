// src/app/core/services/notification.service.ts
import { Injectable, Injector, runInInjectionContext } from '@angular/core';
import {
  Firestore, collection, addDoc, query, where, getDocs,
  orderBy, doc, updateDoc, deleteDoc, Timestamp, writeBatch,
  getDoc, DocumentSnapshot
} from '@angular/fire/firestore';
import { AuthService } from './auth.service';
import { Router } from '@angular/router';

export interface Notification {
  id?: string;
  type: 'table_submitted' | 'table_reviewed_by_production' | 'requisition_confirmed' | 'requisition_removed' | 'requisition_scheduled';
  tableId: string;
  tableName: string;
  submittedBy: string;
  submittedByName?: string;
  submittedAt: any;
  read: boolean;
  readAt?: any;
  createdAt: any;
  userId?: string;
}

@Injectable({ providedIn: 'root' })
export class NotificationService {

  constructor(
    private firestore: Firestore,
    private auth: AuthService,
    private injector: Injector,
    private router: Router
  ) {}

  private run<T>(fn: () => Promise<T>): Promise<T> {
    return runInInjectionContext(this.injector, fn);
  }

  /**
   * Send notification to all production users when a table is submitted
   */
  async sendTableSubmittedNotification(tableId: string, tableName: string, submittedBy: string): Promise<void> {
    try {
      console.log('Sending table submitted notification for table:', tableId);

      // Get all production users
      const usersRef = collection(this.firestore, 'users');
      const productionUsersQuery = query(usersRef, where('role', '==', 'production'));
      const productionUsersSnapshot = await this.run(() => getDocs(productionUsersQuery));

      if (productionUsersSnapshot.empty) {
        console.log('No production users found to notify');
        return;
      }

      // Get the name of the user who submitted the table
      const userDocRef = doc(this.firestore, 'users', submittedBy);
      const userDoc = await this.run(() => getDoc(userDocRef));
      let submittedByName = 'A user';
      if (userDoc.exists()) {
        const userData = userDoc.data() as any;
        submittedByName = userData['email'] || userData['full_name'] || 'A user';
      }

      // Create notification for each production user
      const notificationsRef = collection(this.firestore, 'notifications');
      const now = Timestamp.now();

      const promises: Promise<any>[] = [];
      
      productionUsersSnapshot.forEach((userDoc) => {
        const notification: Omit<Notification, 'id'> = {
          type: 'table_submitted',
          tableId,
          tableName,
          submittedBy,
          submittedByName,
          submittedAt: now,
          read: false,
          createdAt: now,
          userId: userDoc.id
        };

        promises.push(
          this.run(() => addDoc(notificationsRef, notification))
        );
      });

      await Promise.all(promises);
      console.log(`Notifications sent to ${productionUsersSnapshot.size} production users`);
    } catch (err) {
      console.error('Failed to send table submitted notification:', err);
    }
  }

  /**
   * Send notification to all procurement users when production submits reviewed table
   */
  async sendTableReviewedByProductionNotification(tableId: string, tableName: string, reviewedBy: string): Promise<void> {
    try {
      const usersRef = collection(this.firestore, 'users');
      const procurementUsersQuery = query(usersRef, where('role', '==', 'procurement'));
      const procurementUsersSnapshot = await this.run(() => getDocs(procurementUsersQuery));

      if (procurementUsersSnapshot.empty) {
        console.log('No procurement users found to notify');
        return;
      }

      const userDocRef = doc(this.firestore, 'users', reviewedBy);
      const userDoc = await this.run(() => getDoc(userDocRef));
      let reviewedByName = 'Production';
      if (userDoc.exists()) {
        const userData = userDoc.data() as any;
        reviewedByName = userData['email'] || userData['full_name'] || 'Production';
      }

      const notificationsRef = collection(this.firestore, 'notifications');
      const now = Timestamp.now();
      const promises: Promise<any>[] = [];

      procurementUsersSnapshot.forEach((userDoc) => {
        const notification: Omit<Notification, 'id'> = {
          type: 'table_reviewed_by_production',
          tableId,
          tableName,
          submittedBy: reviewedBy,
          submittedByName: reviewedByName,
          submittedAt: now,
          read: false,
          createdAt: now,
          userId: userDoc.id
        };
        promises.push(this.run(() => addDoc(notificationsRef, notification)));
      });

      await Promise.all(promises);
      console.log(`Notifications sent to ${procurementUsersSnapshot.size} procurement users`);
    } catch (err) {
      console.error('Failed to send table reviewed notification:', err);
    }
  }

  /**
   * Get unread notifications for the current user
   */
  async getUnreadNotifications(): Promise<Notification[]> {
    try {
      const currentUser = await this.auth.getCurrentUserPromise();
      if (!currentUser) return [];

      const notificationsRef = collection(this.firestore, 'notifications');
      const q = query(
        notificationsRef,
        where('userId', '==', currentUser.uid),
        where('read', '==', false),
        orderBy('createdAt', 'desc')
      );

      const snapshot = await this.run(() => getDocs(q));
      
      const notifications: Notification[] = [];
      snapshot.forEach(doc => {
        const data = doc.data() as Omit<Notification, 'id'>;
        notifications.push({ id: doc.id, ...data } as Notification);
      });

      return notifications;
    } catch (err) {
      console.error('Failed to get unread notifications:', err);
      return [];
    }
  }

  /**
   * Get all notifications for the current user
   */
  async getAllNotifications(): Promise<Notification[]> {
    try {
      const currentUser = await this.auth.getCurrentUserPromise();
      if (!currentUser) return [];

      const notificationsRef = collection(this.firestore, 'notifications');
      const q = query(
        notificationsRef,
        where('userId', '==', currentUser.uid),
        orderBy('createdAt', 'desc')
      );

      const snapshot = await this.run(() => getDocs(q));
      
      const notifications: Notification[] = [];
      snapshot.forEach(doc => {
        const data = doc.data() as Omit<Notification, 'id'>;
        notifications.push({ id: doc.id, ...data } as Notification);
      });

      return notifications;
    } catch (err) {
      console.error('Failed to get all notifications:', err);
      return [];
    }
  }

  /**
   * Mark a notification as read
   */
  async markAsRead(notificationId: string): Promise<boolean> {
    try {
      const notificationRef = doc(this.firestore, 'notifications', notificationId);
      await this.run(() => updateDoc(notificationRef, {
        read: true,
        readAt: Timestamp.now()
      }));
      return true;
    } catch (err) {
      console.error('Failed to mark notification as read:', err);
      return false;
    }
  }

  /**
   * Mark all notifications as read for the current user
   */
  async markAllAsRead(): Promise<boolean> {
    try {
      const currentUser = await this.auth.getCurrentUserPromise();
      if (!currentUser) return false;

      const notificationsRef = collection(this.firestore, 'notifications');
      const q = query(
        notificationsRef,
        where('userId', '==', currentUser.uid),
        where('read', '==', false)
      );

      const snapshot = await this.run(() => getDocs(q));
      
      if (snapshot.empty) {
        return true;
      }
      
      const batch = writeBatch(this.firestore);

      snapshot.forEach(doc => {
        batch.update(doc.ref, {
          read: true,
          readAt: Timestamp.now()
        });
      });

      await batch.commit();
      return true;
    } catch (err) {
      console.error('Failed to mark all notifications as read:', err);
      return false;
    }
  }

  /**
   * Delete a notification
   */
  async deleteNotification(notificationId: string): Promise<boolean> {
    try {
      const notificationRef = doc(this.firestore, 'notifications', notificationId);
      await this.run(() => deleteDoc(notificationRef));
      return true;
    } catch (err) {
      console.error('Failed to delete notification:', err);
      return false;
    }
  }

  /**
   * Subscribe to real-time notifications (for the header bell icon)
   */
  subscribeToNotifications(callback: (notifications: Notification[]) => void): () => void {
    let unsubscribe: (() => void) | null = null;
    
    // Use setTimeout to avoid blocking
    setTimeout(() => {
      runInInjectionContext(this.injector, async () => {
        try {
          const currentUser = await this.auth.getCurrentUserPromise();
          if (!currentUser) {
            console.log('No current user for notification subscription');
            return;
          }

          console.log('Setting up notification subscription for user:', currentUser.uid);

          const notificationsRef = collection(this.firestore, 'notifications');
          const q = query(
            notificationsRef,
            where('userId', '==', currentUser.uid),
            where('read', '==', false),
            orderBy('createdAt', 'desc')
          );

          const { onSnapshot } = await import('@angular/fire/firestore');
          
          unsubscribe = onSnapshot(q, (snapshot) => {
            console.log('Notification snapshot received, size:', snapshot.size);
            const notifications: Notification[] = [];
            snapshot.forEach(doc => {
              const data = doc.data() as Omit<Notification, 'id'>;
              notifications.push({ id: doc.id, ...data } as Notification);
            });
            callback(notifications);
          }, (error) => {
            console.error('Error in notification snapshot:', error);
          });
        } catch (err) {
          console.error('Failed to subscribe to notifications:', err);
        }
      });
    }, 0);

    return () => {
      console.log('Unsubscribing from notifications');
      if (unsubscribe) {
        unsubscribe();
      }
    };
  }
}