// src/core/services/database.service.ts
import { Injectable } from '@angular/core';
import { Firestore, collection, addDoc, getDocs, query, where, doc, getDoc, updateDoc, deleteDoc, writeBatch, DocumentData, DocumentReference, QuerySnapshot, DocumentSnapshot } from '@angular/fire/firestore';
import { AuthService } from './auth.service';
import { firstValueFrom } from 'rxjs';
import { User } from '../models/database.model';
import * as XLSX from 'xlsx';

interface MasterData {
  sku_code: string;
  sku_name: string;
  qty_per_unit?: number | null;
  unit?: string;
  qty_per_pack?: number | null;
  pack_unit?: string;
  projected_yield_per_batch?: number | null;
  yield_unit?: string;
  category: string;
  raw_material: string;
  qty_per_batch?: number | null;
  batch_unit?: string;
  type?: string;
  supplier?: string | null;
  created_at?: Date;
  updated_at?: Date;
}

interface InventoryItem {
  id?: string;
  sku_code: string;
  sku_name: string;
  category: string;
  supplier: string;
  qty: number;
  created_at?: Date;
  updated_at?: Date;
}

@Injectable({ providedIn: 'root' })
export class DatabaseService {
  constructor(
    private firestore: Firestore,
    private auth: AuthService
  ) {}

  // ────────────────────────────────────────────────
  //  User
  // ────────────────────────────────────────────────

  async getCurrentUser(): Promise<User | null> {
    const authUser = await firstValueFrom(this.auth.getCurrentUser());
    if (!authUser) return null;

    try {
      const userDoc = await getDoc(doc(this.firestore, 'users', authUser.uid));
      if (userDoc.exists()) {
        const data = userDoc.data() as any;
        return {
          id: authUser.uid,
          email: data.email || authUser.email || undefined,
          full_name: data.full_name || data.name || undefined,
          username: data.username || undefined,
          role: data.role || undefined
        };
      }
      return { id: authUser.uid, email: authUser.email || undefined } as User;
    } catch (err) {
      console.error('getCurrentUser failed', err);
      return { id: authUser?.uid, email: authUser?.email } as User;
    }
  }

  // ────────────────────────────────────────────────
  //  Master Data Upload & Queries
  // ────────────────────────────────────────────────

  async uploadMasterData(file: File): Promise<{ success: boolean; count?: number; error?: string }> {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      // Get raw rows (array of arrays)
      const json: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', blankrows: false });

      // Skip header row
      const dataRows = json.slice(1);

      const batch = writeBatch(this.firestore);
      const colRef = collection(this.firestore, 'masterData');

      let savedCount = 0;

      for (const row of dataRows) {
        // Skip empty or invalid rows
        if (!Array.isArray(row) || row.length < 5 || !row[0]?.toString().trim()) {
          continue;
        }

        const docData: MasterData = {
          sku_code: String(row[0] || '').trim(),
          sku_name: String(row[1] || '').trim(),
          qty_per_unit: row[2] ? Number(row[2]) : null,
          unit: String(row[3] || '').trim(),
          qty_per_pack: row[4] ? Number(row[4]) : null,
          pack_unit: String(row[5] || '').trim(),
          projected_yield_per_batch: row[6] ? Number(row[6]) : null,
          yield_unit: String(row[7] || '').trim(),
          category: String(row[8] || '').trim(),
          raw_material: String(row[9] || '').trim(),
          qty_per_batch: row[10] ? Number(row[10]) : null,
          batch_unit: String(row[11] || '').trim(),
          type: String(row[12] || '').trim(),
          supplier: row[13] ? String(row[13]).trim() : null,
          created_at: new Date(),
          updated_at: new Date()
        };

        const skuCode = docData.sku_code as string;
        const rawMaterial = docData.raw_material || 'no-material';

        const docId = `${skuCode}_${rawMaterial}`
          .replace(/[^a-zA-Z0-9_-]/g, '_')
          .substring(0, 1500);

        const docRef = doc(colRef, docId);
        batch.set(docRef, docData, { merge: true });
        savedCount++;
      }

      await batch.commit();
      return { success: true, count: savedCount };
    } catch (err: any) {
      console.error('Master data upload failed', err);
      return { success: false, error: err.message || 'Unknown error' };
    }
  }

  async getUniqueCategories(): Promise<string[]> {
    const snapshot = await getDocs(collection(this.firestore, 'masterData'));
    const cats = new Set<string>();
    
    snapshot.forEach(doc => {
      const data = doc.data() as MasterData;
      const category = (data?.category || '').trim();
      if (category) cats.add(category);
    });
    
    return Array.from(cats).sort();
  }

  async getSkusByCategory(category: string): Promise<{ sku_code: string; sku_name: string }[]> {
    if (!category?.trim()) return [];

    const q = query(
      collection(this.firestore, 'masterData'),
      where('category', '==', category)
    );
    
    const snapshot = await getDocs(q);
    const map = new Map<string, string>();
    
    snapshot.forEach(doc => {
      const data = doc.data() as MasterData;
      const code = (data.sku_code || '').trim();
      const name = (data.sku_name || '').trim();
      if (code && name) map.set(code, name);
    });

    return Array.from(map, ([sku_code, sku_name]) => ({ sku_code, sku_name }));
  }

  async getMaterialsForSku(skuCode: string): Promise<any[]> {
    if (!skuCode?.trim()) return [];

    const q = query(
      collection(this.firestore, 'masterData'),
      where('sku_code', '==', skuCode)
    );
    
    const snapshot = await getDocs(q);
    const materials: any[] = [];
    
    snapshot.forEach(doc => {
      const data = doc.data() as MasterData;
      const material = {
        raw_material: data.raw_material || '',
        quantity_per_batch: data.qty_per_batch ?? null,
        unit: data.batch_unit || '',
        type: data.type || ''
      };
      
      if (material.raw_material.trim() !== '') {
        materials.push(material);
      }
    });

    return materials;
  }

  // ────────────────────────────────────────────────
  //  Inventory (user-managed stock)
  // ────────────────────────────────────────────────

  async addInventoryItem(item: any): Promise<{ success: boolean; id?: string }> {
    try {
      const docRef = await addDoc(collection(this.firestore, 'inventory'), {
        ...item,
        created_at: new Date(),
        updated_at: new Date()
      });
      return { success: true, id: docRef.id };
    } catch (err) {
      console.error('addInventoryItem failed', err);
      return { success: false };
    }
  }

  async getInventoryItems(): Promise<any[]> {
    try {
      const snapshot = await getDocs(collection(this.firestore, 'inventory'));
      return snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
    } catch (err) {
      console.error('getInventoryItems failed', err);
      return [];
    }
  }

  // ────────────────────────────────────────────────
  //  Tables + Requisitions
  // ────────────────────────────────────────────────

  async getUserTables(userId: string): Promise<any[]> {
    try {
      const q = query(
        collection(this.firestore, 'tables'),
        where('user_id', '==', userId)
      );
      const snapshot = await getDocs(q);
      return snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
    } catch (err) {
      console.error('getUserTables failed', err);
      return [];
    }
  }

  async createUserTable(data: any): Promise<{ success: boolean; tableId?: string }> {
    try {
      const docRef = await addDoc(collection(this.firestore, 'tables'), data);
      return { success: true, tableId: docRef.id };
    } catch (err) {
      console.error('createUserTable failed', err);
      return { success: false };
    }
  }

  async updateTableName(tableId: string, name: string): Promise<boolean> {
    try {
      await updateDoc(doc(this.firestore, 'tables', tableId), {
        name,
        updated_at: new Date().toISOString()
      });
      return true;
    } catch (err) {
      console.error('updateTableName failed', err);
      return false;
    }
  }

  async deleteTable(tableId: string): Promise<boolean> {
    try {
      // Get all requisitions for this table
      const q = query(
        collection(this.firestore, 'requisitions'),
        where('table_id', '==', tableId)
      );
      const requisitionsSnapshot = await getDocs(q);

      const batch = writeBatch(this.firestore);
      
      // Delete all requisitions
      requisitionsSnapshot.forEach(doc => {
        batch.delete(doc.ref);
      });
      
      // Delete the table itself
      batch.delete(doc(this.firestore, 'tables', tableId));
      
      await batch.commit();
      return true;
    } catch (err) {
      console.error('deleteTable failed', err);
      return false;
    }
  }

  async getTableRequisitions(tableId: string): Promise<any[]> {
    try {
      const q = query(
        collection(this.firestore, 'requisitions'),
        where('table_id', '==', tableId)
      );
      const snapshot = await getDocs(q);
      return snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
    } catch (err) {
      console.error('getTableRequisitions failed', err);
      return [];
    }
  }

  async createRequisition(data: any, materials: any[]): Promise<{ success: boolean; id?: string }> {
    try {
      const docRef = await addDoc(collection(this.firestore, 'requisitions'), {
        ...data,
        materials,
        created_at: new Date().toISOString(),
        updated_at: new Date().toISOString()
      });
      return { success: true, id: docRef.id };
    } catch (err) {
      console.error('createRequisition failed', err);
      return { success: false };
    }
  }

  async updateRequisitionQty(id: string, qty: number): Promise<boolean> {
    try {
      await updateDoc(doc(this.firestore, 'requisitions', id), {
        qty_needed: qty,
        updated_at: new Date().toISOString()
      });
      return true;
    } catch (err) {
      console.error('updateRequisitionQty failed', err);
      return false;
    }
  }

  async updateRequisitionSupplier(id: string, supplier: string): Promise<boolean> {
    try {
      await updateDoc(doc(this.firestore, 'requisitions', id), {
        supplier,
        updated_at: new Date().toISOString()
      });
      return true;
    } catch (err) {
      console.error('updateRequisitionSupplier failed', err);
      return false;
    }
  }

  async deleteRequisition(id: string): Promise<boolean> {
    try {
      await deleteDoc(doc(this.firestore, 'requisitions', id));
      return true;
    } catch (err) {
      console.error('deleteRequisition failed', err);
      return false;
    }
  }

  async updateTableItemCount(tableId: string, count: number): Promise<boolean> {
    try {
      await updateDoc(doc(this.firestore, 'tables', tableId), {
        item_count: count,
        updated_at: new Date().toISOString()
      });
      return true;
    } catch (err) {
      console.error('updateTableItemCount failed', err);
      return false;
    }
  }
  
}