import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterLink } from '@angular/router';
import { DatabaseService } from '../../../core/services/database.service';
import { AuthService } from '../../../core/services/auth.service';
import { Firestore, collection, doc, getDoc, getDocs, query, where } from '@angular/fire/firestore';

interface RecentRequisition {
  id: string;
  reqNumber: string;
  skuName: string;
  status: string;
  date: string;
}

interface StatusCount {
  label: string;
  count: number;
  class: string;
}

interface DayActivity {
  label: string;
  count: number;
  height: number;
}

interface TopMaterial {
  name: string;
  quantity: number;
  unit: string;
}

@Component({
  selector: 'app-page1',
  standalone: true,
  imports: [CommonModule, RouterLink],
  templateUrl: './page1.component.html',
  styleUrl: './page1.component.css',
})
export class Page1Component implements OnInit {
  isLoading = true;
  loadError: string | null = null;
  userRole = '';

  // Production
  productionLineCount = 0;
  productionItemCount = 0;
  totalProductionQty = 0;

  // Ordering
  requisitionTableCount = 0;
  totalRequisitions = 0;
  pendingRequisitions = 0;
  submittedRequisitions = 0;
  deliveredRequisitions = 0;
  submittedTablesCount = 0;

  // Usage
  totalMaterials = 0;
  totalMaterialQty = 0;
  topMaterials: TopMaterial[] = [];

  // Derived
  fulfillmentRate = 0;
  recentRequisitions: RecentRequisition[] = [];
  statusBreakdown: StatusCount[] = [];
  weeklyActivity: DayActivity[] = [];

  userId = '';

  constructor(
    private db: DatabaseService,
    private auth: AuthService,
    private firestore: Firestore
  ) {}

  async ngOnInit() {
    const user = await this.auth.getCurrentUserPromise();
    if (!user) {
      this.isLoading = false;
      return;
    }

    this.userId = user.uid;
    await this.loadUserRole();
    await this.loadDashboardData();
  }

  async refresh() {
    await this.loadDashboardData();
  }

  private async loadUserRole() {
    try {
      const userDoc = await getDoc(doc(this.firestore, 'users', this.userId));
      if (userDoc.exists()) {
        this.userRole = userDoc.data()['role'] || 'user';
      }
    } catch {
      this.userRole = 'user';
    }
  }

  private async loadDashboardData() {
    this.isLoading = true;
    this.loadError = null;

    try {
      const [productionTables, requisitionTables, requisitions] = await Promise.all([
        this.loadProductionTables(),
        this.loadRequisitionTables(),
        this.db.getUserRequisitions(this.userId),
      ]);

      this.productionLineCount = productionTables.length;
      this.requisitionTableCount = requisitionTables.length;
      this.submittedTablesCount = requisitionTables.filter(t => t.submitted).length;

      await this.loadProductionStats(productionTables);
      this.loadOrderingStats(requisitions);
      await this.loadUsageStats(requisitions);
      this.buildWeeklyActivity(requisitions);
      this.buildStatusBreakdown(requisitions);
      this.buildRecentRequisitions(requisitions);
    } catch (err) {
      console.error('Failed to load dashboard data', err);
      this.loadError = 'Failed to load dashboard data. Please try again.';
    } finally {
      this.isLoading = false;
    }
  }

  private async loadProductionTables(): Promise<any[]> {
    const tablesRef = collection(this.firestore, 'tables');
    const q = query(
      tablesRef,
      where('user_id', '==', this.userId),
      where('type', '==', 'production')
    );
    const snapshot = await getDocs(q);
    return snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
  }

  private async loadRequisitionTables(): Promise<any[]> {
    const tablesRef = collection(this.firestore, 'tables');
    const q = query(
      tablesRef,
      where('user_id', '==', this.userId),
      where('type', '==', 'requisition')
    );
    const snapshot = await getDocs(q);
    return snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
  }

  private async loadProductionStats(productionTables: any[]) {
    this.productionItemCount = 0;
    this.totalProductionQty = 0;

    if (productionTables.length === 0) return;

    const requisitionsRef = collection(this.firestore, 'requisitions');
    const snapshots = await Promise.all(
      productionTables.map(table =>
        getDocs(query(
          requisitionsRef,
          where('table_id', '==', table.id),
          where('user_id', '==', this.userId)
        ))
      )
    );

    snapshots.forEach(snapshot => {
      snapshot.forEach(docSnap => {
        const data = docSnap.data();
        this.productionItemCount++;
        this.totalProductionQty += Number(data['qty_needed'] ?? data['quantity'] ?? 0);
      });
    });
  }

  private loadOrderingStats(requisitions: any[]) {
    this.totalRequisitions = requisitions.length;
    this.pendingRequisitions = requisitions.filter(r => r.status === 'Pending').length;
    this.submittedRequisitions = requisitions.filter(r =>
      r.status === 'Submitted' ||
      r.status === 'Production_Confirmed' ||
      r.status === 'Scheduled'
    ).length;
    this.deliveredRequisitions = requisitions.filter(r =>
      r.status === 'Delivered' || r.status === 'Partially_Delivered'
    ).length;

    const completed = this.deliveredRequisitions;
    this.fulfillmentRate = this.totalRequisitions > 0
      ? Math.round((completed / this.totalRequisitions) * 1000) / 10
      : 0;
  }

  private async loadUsageStats(requisitions: any[]) {
    const skuCache = new Map<string, any[]>();
    const materialMap = new Map<string, { quantity: number; unit: string }>();

    const uniqueSkus = [...new Set(
      requisitions
        .map(r => this.db.normalizeSkuCode(r.sku_code || r.skuCode || ''))
        .filter(Boolean)
    )];

    await Promise.all(uniqueSkus.map(async sku => {
      const materials = await this.db.getMaterialsForSku(sku);
      skuCache.set(sku, materials);
    }));

    for (const req of requisitions) {
      const skuCode = this.db.normalizeSkuCode(req.sku_code || req.skuCode || '');
      if (!skuCode) continue;

      const materials = skuCache.get(skuCode) || [];
      const qtyRequired = Number(req.qty_needed ?? req.quantity ?? 0);

      for (const mat of materials) {
        if (!mat.raw_material) continue;
        const total = Number(mat.quantity_per_batch ?? 0) * qtyRequired;
        const key = mat.raw_material;
        const existing = materialMap.get(key);
        if (existing) {
          existing.quantity += total;
        } else {
          materialMap.set(key, { quantity: total, unit: mat.unit || '' });
        }
      }
    }

    this.totalMaterials = materialMap.size;
    this.totalMaterialQty = [...materialMap.values()].reduce((sum, m) => sum + m.quantity, 0);

    this.topMaterials = [...materialMap.entries()]
      .map(([name, data]) => ({ name, quantity: data.quantity, unit: data.unit }))
      .sort((a, b) => b.quantity - a.quantity)
      .slice(0, 5);
  }

  private buildWeeklyActivity(requisitions: any[]) {
    const days: DayActivity[] = [];
    const now = new Date();

    for (let i = 6; i >= 0; i--) {
      const date = new Date(now);
      date.setDate(now.getDate() - i);
      date.setHours(0, 0, 0, 0);

      const nextDay = new Date(date);
      nextDay.setDate(date.getDate() + 1);

      const count = requisitions.filter(r => {
        const raw = r.created_at || r.submitted_at;
        if (!raw) return false;
        const created = new Date(raw);
        return created >= date && created < nextDay;
      }).length;

      days.push({
        label: date.toLocaleDateString('en-US', { weekday: 'short' }),
        count,
        height: 0,
      });
    }

    const maxCount = Math.max(...days.map(d => d.count), 1);
    days.forEach(d => {
      d.height = Math.max(8, Math.round((d.count / maxCount) * 100));
    });

    this.weeklyActivity = days;
  }

  private buildStatusBreakdown(requisitions: any[]) {
    const statusMap = new Map<string, number>();

    for (const req of requisitions) {
      const status = req.status || 'Pending';
      statusMap.set(status, (statusMap.get(status) || 0) + 1);
    }

    const classMap: Record<string, string> = {
      Pending: 'badge-pending',
      Submitted: 'badge-processing',
      Production_Confirmed: 'badge-processing',
      Scheduled: 'badge-processing',
      Delivered: 'badge-success',
      Partially_Delivered: 'badge-pending',
      Removed: 'badge-processing',
      Approved: 'badge-success',
      Rejected: 'badge-processing',
    };

    this.statusBreakdown = [...statusMap.entries()]
      .map(([label, count]) => ({
        label: label.replace(/_/g, ' '),
        count,
        class: classMap[label] || 'badge-pending',
      }))
      .sort((a, b) => b.count - a.count);
  }

  private buildRecentRequisitions(requisitions: any[]) {
    this.recentRequisitions = [...requisitions]
      .sort((a, b) => {
        const dateA = new Date(a.created_at || a.submitted_at || 0).getTime();
        const dateB = new Date(b.created_at || b.submitted_at || 0).getTime();
        return dateB - dateA;
      })
      .slice(0, 8)
      .map(r => ({
        id: r.id,
        reqNumber: r.reqNumber || r.req_number || `REQ-${(r.id || '').slice(0, 6)}`,
        skuName: r.skuName || r.sku_name || r.skuCode || r.sku_code || 'Unknown',
        status: (r.status || 'Pending').replace(/_/g, ' '),
        date: this.formatDate(r.created_at || r.submitted_at),
      }));
  }

  private formatDate(value: any): string {
    if (!value) return '—';
    try {
      return new Date(value).toLocaleDateString('en-US', {
        month: 'short',
        day: 'numeric',
        year: 'numeric',
      });
    } catch {
      return '—';
    }
  }

  getStatusBadgeClass(status: string): string {
    const normalized = status.toLowerCase().replace(/\s/g, '_');
    switch (normalized) {
      case 'delivered':
      case 'approved':
        return 'badge-success';
      case 'submitted':
      case 'production_confirmed':
      case 'scheduled':
      case 'removed':
        return 'badge-processing';
      default:
        return 'badge-pending';
    }
  }

  get weeklyActivityTotal(): number {
    return this.weeklyActivity.reduce((sum, d) => sum + d.count, 0);
  }
}
