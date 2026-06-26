import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterLink } from '@angular/router';
import { DatabaseService } from '../../../core/services/database.service';
import { AuthService } from '../../../core/services/auth.service';
import {
  ForecastMaterial,
  ForecastSku,
  OrderingForecastService,
} from '../../../core/services/ordering-forecast.service';
import { Firestore, collection, doc, getDoc, getDocs, query, where } from '@angular/fire/firestore';

interface RecentRequisition {
  id: string;
  reqNumber: string;
  tableName: string;
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
  readonly skeletonRows = [1, 2, 3, 4, 5];
  readonly skeletonStats = [1, 2, 3, 4];

  // Monthly overview
  topItemName = '';
  topItemCode = '';
  topRawMaterial = '';
  monthlyRequisitionSlips = 0;
  monthlyPendingInProgress = 0;

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
  activityDays = 7;
  readonly activityDayOptions = [7, 15, 30];

  // Forecast
  forecastDays = 30;
  readonly forecastDayOptions = [7, 30];
  forecastSkus: ForecastSku[] = [];
  forecastMaterials: ForecastMaterial[] = [];
  forecastHasEnoughData = false;
  forecastHasHistory = false;

  userId = '';
  private allRequisitions: any[] = [];
  private tableNameMap = new Map<string, string>();
  private skuCache = new Map<string, any[]>();

  constructor(
    private db: DatabaseService,
    private auth: AuthService,
    private firestore: Firestore,
    private forecastService: OrderingForecastService
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

  setActivityDays(days: number) {
    if (this.activityDays === days) return;
    this.activityDays = days;
    if (this.isLoading) {
      this.weeklyActivity = this.buildSkeletonActivity(days);
      return;
    }
    this.buildActivityChart(this.allRequisitions, days);
  }

  setForecastDays(days: number) {
    if (this.forecastDays === days) return;
    this.forecastDays = days;
    if (this.isLoading) return;
    this.buildForecast(this.allRequisitions);
  }

  private buildSkeletonActivity(days: number): DayActivity[] {
    const now = new Date();
    return Array.from({ length: days }, (_, index) => {
      const date = new Date(now);
      date.setDate(now.getDate() - (days - 1 - index));
      const label = days <= 7
        ? date.toLocaleDateString('en-US', { weekday: 'short' })
        : date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
      return { label, count: 0, height: 18 + (index % 4) * 12 };
    });
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

  /** Admin and production see organization-wide overview stats. */
  private hasOverallAccess(): boolean {
    return this.userRole === 'admin' || this.userRole === 'production';
  }

  private async loadDashboardData() {
    this.isLoading = true;
    this.loadError = null;
    this.weeklyActivity = this.buildSkeletonActivity(this.activityDays);

    try {
      const [productionTables, requisitionTables, requisitions] = await Promise.all([
        this.loadProductionTables(),
        this.loadRequisitionTables(),
        this.db.getRequisitionsForDashboard(this.userId, this.userRole),
      ]);

      this.allRequisitions = requisitions;
      await this.loadTableNameMap(requisitions);

      this.productionLineCount = productionTables.length;
      this.requisitionTableCount = requisitionTables.length;
      this.submittedTablesCount = requisitionTables.filter(t => t.submitted).length;

      const monthlyRequisitions = this.filterMonthlyRequisitions(requisitions);

      await this.loadProductionStats(productionTables);
      this.loadOrderingStats(requisitions);
      this.skuCache = await this.buildSkuCache(requisitions);
      this.loadUsageStats(requisitions, this.skuCache);
      this.buildForecast(requisitions);
      await this.buildMonthlyOverview(monthlyRequisitions);
      this.buildActivityChart(requisitions, this.activityDays);
      this.buildStatusBreakdown(requisitions);
      this.buildRecentRequisitions(requisitions);
    } catch (err) {
      console.error('Failed to load dashboard data', err);
      this.loadError = 'Failed to load dashboard data. Please try again.';
    } finally {
      this.isLoading = false;
    }
  }

  private async loadTableNameMap(requisitions: any[]) {
    this.tableNameMap.clear();
    const tableIds = [...new Set(
      requisitions.map(r => r.table_id).filter((id): id is string => Boolean(id))
    )];

    await Promise.all(tableIds.map(async tableId => {
      try {
        const tableDoc = await getDoc(doc(this.firestore, 'tables', tableId));
        if (tableDoc.exists()) {
          this.tableNameMap.set(tableId, tableDoc.data()['name'] || 'Untitled');
        }
      } catch {
        // ignore missing table lookups
      }
    }));
  }

  private getRequisitionDate(req: any): Date | null {
    const raw = req.created_at || req.submitted_at;
    if (!raw) return null;
    const date = new Date(raw);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  private filterMonthlyRequisitions(requisitions: any[]): any[] {
    const now = new Date();
    const month = now.getMonth();
    const year = now.getFullYear();

    return requisitions.filter(req => {
      const date = this.getRequisitionDate(req);
      return date !== null && date.getMonth() === month && date.getFullYear() === year;
    });
  }

  private async buildMonthlyOverview(monthlyRequisitions: any[]) {
    this.monthlyRequisitionSlips = monthlyRequisitions.length;

    const inProgressStatuses = new Set([
      'Pending',
      'Submitted',
      'Production_Confirmed',
      'Scheduled',
    ]);
    this.monthlyPendingInProgress = monthlyRequisitions.filter(r =>
      inProgressStatuses.has(r.status || 'Pending')
    ).length;

    const skuCounts = new Map<string, { name: string; count: number }>();
    for (const req of monthlyRequisitions) {
      const code = this.db.normalizeSkuCode(req.sku_code || req.skuCode || '');
      if (!code) continue;

      const name = String(req.skuName || req.sku_name || code).trim();
      const existing = skuCounts.get(code);
      if (existing) {
        existing.count += 1;
      } else {
        skuCounts.set(code, { name, count: 1 });
      }
    }

    const topSku = [...skuCounts.entries()].sort((a, b) => b[1].count - a[1].count)[0];
    if (topSku) {
      this.topItemCode = topSku[0];
      this.topItemName = topSku[1].name;
    } else {
      this.topItemCode = '';
      this.topItemName = '';
    }

    this.topRawMaterial = monthlyRequisitions.length > 0
      ? await this.computeTopRawMaterial(monthlyRequisitions)
      : '';
  }

  private async computeTopRawMaterial(monthlyRequisitions: any[]): Promise<string> {
    const skuCache = new Map<string, any[]>();
    const materialTotals = new Map<string, number>();

    const uniqueSkus = [...new Set(
      monthlyRequisitions
        .map(r => this.db.normalizeSkuCode(r.sku_code || r.skuCode || ''))
        .filter(Boolean)
    )];

    await Promise.all(uniqueSkus.map(async sku => {
      const materials = await this.db.getMaterialsForSku(sku);
      skuCache.set(sku, materials);
    }));

    for (const req of monthlyRequisitions) {
      const skuCode = this.db.normalizeSkuCode(req.sku_code || req.skuCode || '');
      if (!skuCode) continue;

      const materials = skuCache.get(skuCode) || [];
      const qtyRequired = Number(req.qty_needed ?? req.quantity ?? 0);

      for (const mat of materials) {
        if (!mat.raw_material) continue;
        const total = Number(mat.quantity_per_batch ?? 0) * qtyRequired;
        const key = mat.raw_material;
        materialTotals.set(key, (materialTotals.get(key) || 0) + total);
      }
    }

    const top = [...materialTotals.entries()].sort((a, b) => b[1] - a[1])[0];
    return top?.[0] || '';
  }

  private async loadProductionTables(): Promise<any[]> {
    const tablesRef = collection(this.firestore, 'tables');
    const q = this.hasOverallAccess()
      ? query(tablesRef, where('type', '==', 'production'))
      : query(tablesRef, where('user_id', '==', this.userId), where('type', '==', 'production'));
    const snapshot = await getDocs(q);
    return snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
  }

  private async loadRequisitionTables(): Promise<any[]> {
    const tablesRef = collection(this.firestore, 'tables');
    const q = this.hasOverallAccess()
      ? query(tablesRef, where('type', '==', 'requisition'))
      : query(tablesRef, where('user_id', '==', this.userId), where('type', '==', 'requisition'));
    const snapshot = await getDocs(q);
    return snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
  }

  private async loadProductionStats(productionTables: any[]) {
    this.productionItemCount = 0;
    this.totalProductionQty = 0;

    if (productionTables.length === 0) return;

    const requisitionsRef = collection(this.firestore, 'requisitions');
    const snapshots = await Promise.all(
      productionTables.map(table => {
        const constraints = [where('table_id', '==', table.id)];
        if (!this.hasOverallAccess()) {
          constraints.push(where('user_id', '==', this.userId));
        }
        return getDocs(query(requisitionsRef, ...constraints));
      })
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

  private async buildSkuCache(requisitions: any[]): Promise<Map<string, any[]>> {
    const skuCache = new Map<string, any[]>();
    const uniqueSkus = [...new Set(
      requisitions
        .map(r => this.db.normalizeSkuCode(r.sku_code || r.skuCode || ''))
        .filter(Boolean)
    )];

    await Promise.all(uniqueSkus.map(async sku => {
      const materials = await this.db.getMaterialsForSku(sku);
      skuCache.set(sku, materials);
    }));

    return skuCache;
  }

  private loadUsageStats(requisitions: any[], skuCache: Map<string, any[]>) {
    const materialMap = new Map<string, { quantity: number; unit: string }>();

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

  private buildForecast(requisitions: any[]) {
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 90);
    this.forecastHasHistory = requisitions.some(req => {
      if ((req.status || '') === 'Removed') return false;
      const date = this.getRequisitionDate(req);
      return date !== null && date >= cutoff;
    });

    const result = this.forecastService.computeForecast(
      requisitions,
      this.skuCache,
      this.forecastDays,
      req => this.getRequisitionDate(req)
    );

    this.forecastSkus = result.topSkus;
    this.forecastMaterials = result.topMaterials;
    this.forecastHasEnoughData = result.hasEnoughData;
  }

  private buildActivityChart(requisitions: any[], days: number) {
    const activity: DayActivity[] = [];
    const now = new Date();

    for (let i = days - 1; i >= 0; i--) {
      const date = new Date(now);
      date.setDate(now.getDate() - i);
      date.setHours(0, 0, 0, 0);

      const nextDay = new Date(date);
      nextDay.setDate(date.getDate() + 1);

      const count = requisitions.filter(r => {
        const created = this.getRequisitionDate(r);
        if (!created) return false;
        return created >= date && created < nextDay;
      }).length;

      const label = days <= 7
        ? date.toLocaleDateString('en-US', { weekday: 'short' })
        : date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });

      activity.push({ label, count, height: 0 });
    }

    const maxCount = Math.max(...activity.map(d => d.count), 1);
    activity.forEach(d => {
      d.height = d.count === 0
        ? 3
        : Math.max(12, Math.round((d.count / maxCount) * 100));
    });

    this.weeklyActivity = activity;
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
        const dateA = this.getRequisitionDate(a)?.getTime() ?? 0;
        const dateB = this.getRequisitionDate(b)?.getTime() ?? 0;
        return dateB - dateA;
      })
      .slice(0, 8)
      .map(r => ({
        id: r.id,
        reqNumber: r.reqNumber || r.req_number || `REQ-${(r.id || '').slice(0, 6)}`,
        tableName: this.tableNameMap.get(r.table_id) || r.table_name || '—',
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

  get currentMonthLabel(): string {
    return new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  }

  get projectedTopSku(): ForecastSku | null {
    return this.forecastSkus[0] ?? null;
  }

  get projectedTopMaterial(): ForecastMaterial | null {
    return this.forecastMaterials[0] ?? null;
  }
}
