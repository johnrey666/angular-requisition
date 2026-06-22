import { Component, OnInit, HostListener, Injector, ChangeDetectorRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClientModule } from '@angular/common/http';
import { runInInjectionContext } from '@angular/core';
import { DatabaseService } from '../../../core/services/database.service';
import { AuthService } from '../../../core/services/auth.service';
import { NotificationService } from '../../../core/services/notification.service';
import { EmailNotificationService } from '../../../core/services/email-notification.service';
import {
  Firestore, doc, collection, query, where, getDocs,
  writeBatch, getDoc, updateDoc, orderBy
} from '@angular/fire/firestore';
import { Router, ActivatedRoute } from '@angular/router';
import * as XLSX from 'xlsx';

interface Material {
  raw_material: string;
  quantity_per_batch: number | null;
  unit: string;
  type: string;
  production_action?: 'confirmed' | 'removed';
}

interface Requisition {
  id: string;
  reqNumber: string;
  type: string;
  dateNeeded?: string;
  skuCode: string;
  skuName: string;
  quantity: number;
  unit: string;
  supplier: string;
  brand?: string;
  status: string;
  category: string;
  created_at?: string;
  user_id?: string;
  user_email?: string;
  table_id?: string;
  table_name?: string;
  submitted_at?: string;
  scheduled_date?: string;
  scheduled_at?: string;
  scheduled_by?: string;
  approved_at?: string;
  approved_by?: string;
  rejection_reason?: string;
  production_action?: 'confirmed' | 'removed';
  production_action_at?: string;
  production_action_by?: string;
  production_action_notes?: string;
  procurement_action?: 'reviewed' | 'pending';
  procurement_action_at?: string;
  procurement_action_by?: string;
  procurement_notes?: string;
  materials?: Material[];
  [key: string]: any;
}

interface Table {
  id: string;
  name: string;
  user_id: string;
  user_email?: string;
  type: 'inventory' | 'requisition' | 'production';
  item_count?: number;
  created_at?: string;
  updated_at?: string;
  submitted?: boolean;
  submitted_at?: string;
  po_file_url?: string;
  po_file_data?: string;
  po_file_mime?: string;
  po_file_name?: string;
  po_file_size?: number;
  po_file_type?: string;
  production_reviewed?: boolean;
  production_reviewed_at?: string;
  production_reviewed_by?: string;
  request_closed?: boolean;
  request_closed_at?: string;
  request_closed_by?: string;
}

interface SkuOption {
  sku_code: string;
  sku_name: string;
}

interface MasterDataRow {
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
}

interface ProcurementMaterialSummary {
  raw_material: string;
  unit: string;
  type: string;
  totalQuantity: number;
  table_id: string;
  table_name: string;
  procurement_action?: 'approved' | 'rejected' | null;
  production_status?: 'confirmed' | 'removed' | null;
}

interface ProcurementTableSummary {
  table_id: string;
  table_name: string;
  uniqueMaterialsCount: number;
  totalRequestedQuantity: number;
  materials: ProcurementMaterialSummary[];
}

// For the raw materials modal (user/production view)
interface RawMaterialModalItem {
  raw_material: string;
  type: string;
  unit: string;
  totalQuantity: number;
  production_status: 'confirmed' | 'removed' | null;
  procurement_action: 'approved' | 'rejected' | null;
}

@Component({
  selector: 'app-page3',
  standalone: true,
  imports: [CommonModule, FormsModule, HttpClientModule],
  templateUrl: './page3.component.html',
  styleUrls: ['./page3.component.css']
})
export class Page3Component implements OnInit {

  categories: string[] = [];
  availableSkus: SkuOption[] = [];

  tables: Table[] = [];
  selectedTableId: string = '';
  selectedTable: Table | null = null;
  showTableDropdown = false;

  requisitions: Requisition[] = [];
  filteredRequisitions: Requisition[] = [];
  paginatedRequisitions: Requisition[] = [];

  productionSubmissions: Requisition[] = [];
  productionReviewed: Requisition[] = [];

  procurementReviewed: Requisition[] = [];
  procurementTableSummaries: ProcurementTableSummary[] = [];
  expandedProcurementOriginalItems: { [tableId: string]: boolean } = {};
  showProcurementOriginalItemsModal = false;
  currentProcurementTableId = '';
  currentProcurementTableName = '';
  expandedProcurementModalRows: { [reqId: string]: boolean } = {};

  expandedRows: { [id: string]: boolean } = {};
  loadingMaterials: { [id: string]: boolean } = {};

  // Raw Materials Modal (user/production)
  showRawMaterialsModal = false;
  rawMaterialsModalLoading = false;
  rawMaterialsModalData: RawMaterialModalItem[] = [];

  // Procurement material actions map: key = `${tableId}|${rawMaterial}` => action
  procurementMaterialActionsMap: { [key: string]: 'approved' | 'rejected' | null } = {};

  showModal = false;
  showTableModal = false;
  showMasterDataModal = false;
  masterDataRows: MasterDataRow[] = [];
  filteredMasterDataRows: MasterDataRow[] = [];
  masterDataSearchQuery = '';
  loadingMasterDataView = false;

  // Collapsible Master Data Properties
  groupedMasterData: { [skuKey: string]: { sku: { code: string; name: string }; materials: MasterDataRow[] } } = {};
  groupedMasterDataArray: Array<{ skuKey: string; sku: { code: string; name: string }; materials: MasterDataRow[]; materialCount: number }> = [];
  filteredGroupedMasterData: Array<{ skuKey: string; sku: { code: string; name: string }; materials: MasterDataRow[]; materialCount: number }> = [];
  expandedSkus: { [skuKey: string]: boolean } = {};

  showScheduleModal = false;
  showApproveModal = false;
  showDeliveryModal = false;
  showMissingNotesModal = false;
  showProductionActionModal = false;
  viewMode: 'my_tables' | 'store_submissions' | 'for_delivery' | 'production_reviewed' | 'procurement_reviewed' = 'my_tables';
  selectedProductionView: 'submissions' | 'reviewed' = 'submissions';
  showAllPending = false;
  submitted = false;
  isLoading = false;
  isSubmitting = false;
  today = new Date().toISOString().split('T')[0];
  tomorrow: string = '';

  // P.O File Upload Properties
  poFile: File | null = null;
  poFileName: string = '';
  poUploadTargetTable: Table | null = null;
  isUploadingPo: boolean = false;
  private readonly maxPoFileBytes = 500 * 1024;

  formData: any = {
    type: '',
    category: '',
    skuName: '',
    quantity: null,
    unit: '',
    dateNeeded: '',
    supplier: '',
    customSupplier: '',
    brand: '',
    customBrand: ''
  };

  selectedRequisition: Requisition | null = null;
  scheduledDate: string = '';
  scheduledTime: string = '';

  approvalNotes: string = '';
  missingMaterialsNotes: string = '';

  productionActionType: 'confirmed' | 'removed' = 'confirmed';
  productionActionNotes: string = '';

  editingRequisition: Requisition | null = null;
  editingTable: Table | null = null;
  newTableName: string = '';
  editTableName: string = '';

  selectedSkuCode: string = '';

  searchQuery = '';
  filterStatus = '';
  currentPage = 1;
  pageSize = 10;
  totalPages = 1;

  showSnackbar = false;
  snackbarMessage = '';
  snackbarType: 'success' | 'error' | 'info' = 'info';
  snackbarTimeout: any;

  importStatus: 'idle' | 'loading' | 'success' | 'error' = 'idle';
  importMessage = '';
  selectedFileName = '';

  userRole: string = '';
  userId: string = '';
  userName: string = '';

  tableNameMap: { [tableId: string]: string } = {};

  private pendingRouteTableId: string | null = null;

  Math = Math;

  constructor(
    private db: DatabaseService,
    private auth: AuthService,
    private firestore: Firestore,
    private router: Router,
    private route: ActivatedRoute,
    private injector: Injector,
    private cdr: ChangeDetectorRef,
    private notificationService: NotificationService,
    private emailNotificationService: EmailNotificationService
  ) {}

  private run<T>(fn: () => Promise<T>): Promise<T> {
    return runInInjectionContext(this.injector, fn);
  }

  async ngOnInit() {
    const todayDate = new Date();
    const tomorrowDate = new Date(todayDate);
    tomorrowDate.setDate(todayDate.getDate() + 1);
    this.tomorrow = tomorrowDate.toISOString().split('T')[0];

    const user = await this.auth.getCurrentUserPromise();

    if (user) {
      this.userId = user.uid;
      this.userName = user.email || '';
      this.pendingRouteTableId = this.route.snapshot.queryParamMap.get('tableId');
      await this.loadUserRole();
      await this.loadCategories();
      await this.loadMasterDataForMaterials();
      this.setViewModeByRole();

      if (this.userRole === 'production') {
        await this.loadTablesDirectly();
        await this.loadProductionSubmissions();
        if (!this.pendingRouteTableId && this.selectedTable) {
          this.filteredRequisitions = this.productionSubmissions.filter(r => r.table_id === this.selectedTable!.id);
        }
      } else if (this.userRole === 'procurement') {
        await this.loadTablesDirectly();
      } else {
        await this.loadTablesDirectly();

        if (this.tables.length > 0 && !this.selectedTable && !this.pendingRouteTableId) {
          const lastTableId = localStorage.getItem(`lastSelectedRequisitionTable_${this.userId}`);
          if (lastTableId && this.tables.some(t => t.id === lastTableId)) {
            this.selectedTable = this.tables.find(t => t.id === lastTableId) || null;
          } else {
            this.selectedTable = this.tables[0];
          }

          if (this.selectedTable) {
            this.selectedTableId = this.selectedTable.id;
            await this.loadRequisitionsDirectly();
          }
        }
      }

      if (this.pendingRouteTableId) {
        await this.applyTableIdFromRoute(this.pendingRouteTableId);
        this.pendingRouteTableId = null;
      }

      this.route.queryParams.subscribe(async params => {
        const tableId = params['tableId'];
        if (tableId) {
          await this.applyTableIdFromRoute(tableId);
        }
      });
    } else {
      this.showToast('Please log in to continue', 'error');
      this.router.navigate(['/login']);
    }
  }

  async loadUserRole() {
    try {
      const userDocRef = doc(this.firestore, 'users', this.userId);
      const userDoc = await this.run(() => getDoc(userDocRef));
      if (userDoc.exists()) {
        const data = userDoc.data() as any;
        this.userRole = data['role'] || 'user';
        this.userName = data['name'] || this.userName;
      } else {
        this.userRole = 'user';
      }
    } catch (err) {
      this.userRole = 'user';
    }
  }

  async loadCategories() {
    try {
      this.categories = await this.db.getUniqueCategories();
    } catch (err) {}
  }

  async loadMasterDataForMaterials() {
    try {
      this.masterDataRows = await this.db.getAllMasterData();
    } catch (err) {
      this.masterDataRows = [];
    }
  }

  setViewModeByRole() {
    if (this.userRole === 'production') {
      this.viewMode = 'store_submissions';
    } else if (this.userRole === 'procurement') {
      this.viewMode = 'for_delivery';
    } else {
      this.viewMode = 'my_tables';
    }
  }

  async loadTablesDirectly() {
    try {
      this.isLoading = true;
      const tablesRef = collection(this.firestore, 'tables');
      let querySnapshot;

      if (this.userRole === 'production') {
        querySnapshot = await this.run(() => {
          const q = query(
            tablesRef,
            where('type', '==', 'requisition'),
            where('submitted', '==', true)
          );
          return getDocs(q);
        });
      } else if (this.userRole === 'procurement') {
        querySnapshot = await this.run(() => {
          const q = query(
            tablesRef,
            where('type', '==', 'requisition'),
            where('submitted', '==', true),
            where('production_reviewed', '==', true)
          );
          return getDocs(q);
        });
      } else {
        querySnapshot = await this.run(() => {
          const q = query(
            tablesRef,
            where('user_id', '==', this.userId),
            where('type', '==', 'requisition')
          );
          return getDocs(q);
        });
      }

      const loadedTables: Table[] = [];
      const userEmailPromises: Promise<void>[] = [];

      querySnapshot.forEach(docSnap => {
        const data = docSnap.data();
        const table: Table = {
          id: docSnap.id,
          name: data['name'] || 'Untitled',
          user_id: data['user_id'] || '',
          type: data['type'] || 'requisition',
          item_count: data['item_count'] || 0,
          submitted: data['submitted'] || false,
          submitted_at: data['submitted_at'],
          created_at: data['created_at'],
          updated_at: data['updated_at'],
          po_file_url: data['po_file_url'],
          po_file_data: data['po_file_data'],
          po_file_mime: data['po_file_mime'],
          po_file_name: data['po_file_name'],
          po_file_size: data['po_file_size'],
          po_file_type: data['po_file_type'],
          production_reviewed: data['production_reviewed'] || false,
          production_reviewed_at: data['production_reviewed_at'],
          production_reviewed_by: data['production_reviewed_by'],
          request_closed: data['request_closed'] || false,
          request_closed_at: data['request_closed_at'],
          request_closed_by: data['request_closed_by']
        };

        if (table.user_id) {
          const emailPromise = this.getUserEmail(table.user_id).then(email => {
            table.user_email = email;
          });
          userEmailPromises.push(emailPromise);
        }

        loadedTables.push(table);
      });

      await Promise.all(userEmailPromises);

      if (this.userRole === 'production') {
        loadedTables.sort((a, b) => {
          if (a.submitted && b.submitted) {
            return (b.submitted_at || '').localeCompare(a.submitted_at || '');
          }
          return (b.created_at || '').localeCompare(a.created_at || '');
        });
      } else if (this.userRole === 'procurement') {
        loadedTables.sort((a, b) => (b.created_at || '').localeCompare(a.created_at || ''));
      }

      this.tables = loadedTables;

      if (this.userRole === 'production') {
        if (this.tables.length > 0 && !this.selectedTable && !this.pendingRouteTableId) {
          this.selectedTable = this.tables[0];
          this.selectedTableId = this.tables[0].id;

          if (this.selectedTable.production_reviewed) {
            this.selectedProductionView = 'reviewed';
            if (this.productionReviewed.length === 0) {
              await this.loadProductionReviewed();
            }
            this.filteredRequisitions = this.productionReviewed.filter(r => r.table_id === this.selectedTableId);
          } else {
            this.selectedProductionView = 'submissions';
            this.filteredRequisitions = this.productionSubmissions.filter(r => r.table_id === this.selectedTableId);
          }

          this.updatePagination();
        }
      } else if (this.userRole === 'procurement') {
        if (this.tables.length > 0 && !this.pendingRouteTableId) {
          this.selectedTable = this.tables[0];
          this.selectedTableId = this.tables[0].id;
        } else if (!this.pendingRouteTableId) {
          this.selectedTable = null;
          this.selectedTableId = '';
        }

        await this.loadProcurementReviewed();

        if (this.selectedTableId) {
          this.filteredRequisitions = this.procurementReviewed.filter(r => r.table_id === this.selectedTableId);
        } else {
          this.filteredRequisitions = [...this.procurementReviewed];
        }
        this.updatePagination();
      } else if (this.tables.length > 0) {
        const lastTableId = localStorage.getItem(`lastSelectedRequisitionTable_${this.userId}`);
        if (lastTableId && this.tables.some(t => t.id === lastTableId)) {
          this.selectedTableId = lastTableId;
          this.selectedTable = this.tables.find(t => t.id === this.selectedTableId) || null;
        } else {
          this.selectedTableId = this.tables[0].id;
          this.selectedTable = this.tables[0];
        }

        if (this.selectedTable) {
          await this.loadRequisitionsDirectly();
        }
      }
    } catch (err) {
      this.showToast('Failed to load tables', 'error');
    } finally {
      this.isLoading = false;
    }
  }

  async loadRequisitionsDirectly() {
    if (!this.selectedTableId) return;

    try {
      this.isLoading = true;

      const querySnapshot = await this.run(() => {
        const requisitionsRef = collection(this.firestore, 'requisitions');
        const q = query(
          requisitionsRef,
          where('table_id', '==', this.selectedTableId),
          where('user_id', '==', this.userId),
          orderBy('created_at', 'desc')
        );
        return getDocs(q);
      });

      const loadedRequisitions: Requisition[] = [];
      querySnapshot.forEach(docSnap => {
        const data = docSnap.data();
        let scheduledDate = null;
        if (data['scheduled_date']) {
          scheduledDate = data['scheduled_date'];
        } else if (data['scheduled_at']) {
          scheduledDate = data['scheduled_at'];
        }
        if (scheduledDate && typeof scheduledDate === 'object' && scheduledDate.toDate) {
          scheduledDate = scheduledDate.toDate().toISOString();
        }
        loadedRequisitions.push({ id: docSnap.id, ...data, scheduled_date: scheduledDate } as Requisition);
      });

      this.requisitions = loadedRequisitions;
      await this.populateMaterialsForAllRequisitions(this.requisitions);
      this.applyFilter();
    } catch (err) {
      this.showToast('Failed to load requisitions', 'error');
    } finally {
      this.isLoading = false;
    }
  }

  private async populateMaterialsForAllRequisitions(reqs: Requisition[]) {
    for (const req of reqs) {
      const skuCode = String(req.skuCode ?? req['sku_code'] ?? '').trim();
      const skuName = String(req.skuName ?? req['sku_name'] ?? '').trim();

      if (skuCode || skuName) {
        const loadedMaterials = this.getMaterialsFromLoadedData(skuCode, skuName);

        if (req.materials && req.materials.length > 0) {
          const existingActions = req.materials.reduce((map: Record<string, Material>, mat) => {
            map[(mat.raw_material || '').toLowerCase()] = mat;
            return map;
          }, {} as Record<string, Material>);

          req.materials = loadedMaterials.map(mat => ({
            ...mat,
            production_action: existingActions[mat.raw_material.toLowerCase()]?.production_action
          }));
        } else {
          req.materials = loadedMaterials;
        }
      }
    }
  }

  async loadProductionSubmissions() {
    this.isLoading = true;
    try {
      const requisitionsRef = collection(this.firestore, 'requisitions');
      const submissionsSnapshot = await this.run(() => {
        const q = query(
          requisitionsRef,
          where('status', '==', 'Submitted'),
          orderBy('submitted_at', 'desc')
        );
        return getDocs(q);
      });

      const submissions: Requisition[] = [];
      submissionsSnapshot.forEach(docSnap => {
        const data = docSnap.data();
        submissions.push({ id: docSnap.id, ...data } as Requisition);
      });

      await this.loadUserEmailsForRequisitions(submissions);
      await this.loadTableNamesForRequisitions(submissions);
      await this.populateMaterialsForAllRequisitions(submissions);
      this.productionSubmissions = submissions;
    } catch (err) {
      this.showToast('Failed to load submissions', 'error');
    } finally {
      this.isLoading = false;
    }
  }

  async loadProductionReviewed() {
    this.isLoading = true;
    try {
      const requisitionsRef = collection(this.firestore, 'requisitions');
      const reviewed: Requisition[] = [];
      const queries: Array<Promise<any>> = [];

      const addQuery = (q: any) => { queries.push(this.run(() => getDocs(q))); };
      addQuery(query(requisitionsRef, where('production_action', '==', 'confirmed')));
      addQuery(query(requisitionsRef, where('production_action', '==', 'removed')));

      const seen = new Set<string>();
      for (const queryPromise of queries) {
        try {
          const snapshot = await queryPromise;
          snapshot.forEach((docSnap: any) => {
            const data = docSnap.data();
            const id = docSnap.id;
            if (!seen.has(id)) {
              seen.add(id);
              reviewed.push({ id, ...data } as Requisition);
            }
          });
        } catch (queryErr) {}
      }

      await this.loadUserEmailsForRequisitions(reviewed);
      await this.loadTableNamesForRequisitions(reviewed);
      await this.populateMaterialsForAllRequisitions(reviewed);
      this.productionReviewed = reviewed;
    } catch (err) {
      this.showToast('Failed to load reviewed items', 'error');
    } finally {
      this.isLoading = false;
    }
  }

  async loadProcurementReviewed() {
    this.isLoading = true;
    try {
      const requisitionsRef = collection(this.firestore, 'requisitions');
      let reviewed: Requisition[] = [];
      const queries: Array<Promise<any>> = [];
      const addQuery = (q: any) => { queries.push(this.run(() => getDocs(q))); };

      try {
        addQuery(query(requisitionsRef, where('production_action', '==', 'confirmed')));
        addQuery(query(requisitionsRef, where('status', '==', 'Partially_Delivered')));
        addQuery(query(requisitionsRef, where('status', '==', 'Delivered')));
        addQuery(query(requisitionsRef, where('status', '==', 'Scheduled')));

        const snapshots = await Promise.all(queries);
        const seen = new Set<string>();

        snapshots.forEach(snapshot => {
          snapshot.forEach((docSnap: any) => {
            const data = docSnap.data();
            const id = docSnap.id;
            if (!seen.has(id)) {
              seen.add(id);
              reviewed.push({ id, ...data } as Requisition);
            }
          });
        });
      } catch (err) {}

      const approvedTableIds = new Set(this.tables.filter(t => t.production_reviewed).map(t => t.id));
      reviewed = reviewed.filter(r => r.table_id && approvedTableIds.has(r.table_id));

      await this.loadUserEmailsForRequisitions(reviewed);
      await this.loadTableNamesForRequisitions(reviewed);
      await this.populateMaterialsForAllRequisitions(reviewed);

      this.procurementReviewed = reviewed;
      this.filteredRequisitions = [...this.procurementReviewed];
      this.computeProcurementTableSummaries();
      this.updatePagination();
    } catch (err) {
      this.showToast('Failed to load reviewed requisitions', 'error');
    } finally {
      this.isLoading = false;
    }
  }

  private async loadTableNamesForRequisitions(reqs: Requisition[]) {
    const tableIds: string[] = [];
    reqs.forEach(r => { if (r.table_id) tableIds.push(r.table_id); });
    const uniqueTableIds = [...new Set(tableIds)];

    const promises = uniqueTableIds.map(async (tid) => {
      if (!this.tableNameMap[tid]) {
        const t = await this.db.getTableById(tid);
        if (t) this.tableNameMap[tid] = t.name || 'Untitled';
      }
    });

    await Promise.all(promises);
  }

  private async loadUserEmailsForRequisitions(reqs: Requisition[]) {
    const userIds: string[] = [];
    reqs.forEach(r => { if (r.user_id) userIds.push(r.user_id); });
    const uniqueUserIds = [...new Set(userIds)];

    const emailPromises = uniqueUserIds.map(async (uid) => {
      try {
        const userDocRef = doc(this.firestore, 'users', uid);
        const userDoc = await this.run(() => getDoc(userDocRef));
        if (userDoc.exists()) {
          const data = userDoc.data();
          const email = data['email'] || 'Unknown';
          reqs.forEach(r => { if (r.user_id === uid) r.user_email = email; });
        }
      } catch (err) {}
    });

    await Promise.all(emailPromises);
  }

  private async getUserEmail(userId: string): Promise<string> {
    try {
      const userDocRef = doc(this.firestore, 'users', userId);
      const userDoc = await this.run(() => getDoc(userDocRef));
      if (userDoc.exists()) {
        const data = userDoc.data();
        return data['email'] || 'Unknown';
      }
    } catch (err) {}
    return 'Unknown';
  }

  getColspan(): number {
    let baseCols = 13;
    if (this.userRole === 'production') baseCols = 14;
    return baseCols;
  }

  async onTableChange() {
    if (!this.selectedTableId) {
      this.requisitions = [];
      this.filteredRequisitions = [];
      this.selectedTable = null;
      return;
    }

    if (this.userRole !== 'production' && this.userRole !== 'procurement') {
      localStorage.setItem(`lastSelectedRequisitionTable_${this.userId}`, this.selectedTableId);
    }

    this.selectedTable = this.tables.find(t => t.id === this.selectedTableId) || null;

    if (this.userRole !== 'production' && this.userRole !== 'procurement') {
      await this.loadRequisitionsDirectly();
    }
  }

  private async applyTableIdFromRoute(tableId: string): Promise<void> {
    if (this.selectedTableId === tableId) return;

    let table = this.tables.find(t => t.id === tableId);
    if (!table) {
      const fetched = await this.fetchTableById(tableId);
      if (fetched) {
        table = fetched;
        if (!this.tables.some(t => t.id === tableId)) {
          this.tables.unshift(table);
        }
      }
    }

    if (table) {
      await this.selectTable(table);
    } else {
      this.showToast('Submitted table could not be found', 'error');
    }
  }

  private async fetchTableById(tableId: string): Promise<Table | null> {
    try {
      const tableDoc = await this.run(() => getDoc(doc(this.firestore, 'tables', tableId)));
      if (!tableDoc.exists()) return null;

      const data = tableDoc.data();
      const table: Table = {
        id: tableDoc.id,
        name: data['name'] || 'Untitled',
        user_id: data['user_id'] || '',
        type: data['type'] || 'requisition',
        item_count: data['item_count'] || 0,
        submitted: data['submitted'] || false,
        submitted_at: data['submitted_at'],
        created_at: data['created_at'],
        updated_at: data['updated_at'],
        po_file_url: data['po_file_url'],
        po_file_data: data['po_file_data'],
        po_file_mime: data['po_file_mime'],
        po_file_name: data['po_file_name'],
        po_file_size: data['po_file_size'],
        po_file_type: data['po_file_type'],
        production_reviewed: data['production_reviewed'] || false,
        production_reviewed_at: data['production_reviewed_at'],
        production_reviewed_by: data['production_reviewed_by'],
        request_closed: data['request_closed'] || false,
        request_closed_at: data['request_closed_at'],
        request_closed_by: data['request_closed_by']
      };

      if (table.user_id) {
        table.user_email = await this.getUserEmail(table.user_id);
      }

      return table;
    } catch {
      return null;
    }
  }

  async selectTable(table: Table) {
    if (this.userRole === 'production') {
      this.selectedTableId = table.id;
      this.selectedTable = table;
      this.showTableDropdown = false;

      if (table.production_reviewed) {
        this.selectedProductionView = 'reviewed';
        if (this.productionReviewed.length === 0) {
          await this.loadProductionReviewed();
        }
        this.filteredRequisitions = this.productionReviewed.filter(r => r.table_id === table.id);
        this.showToast(`Showing reviewed items from table: ${table.name}`, 'info');
      } else {
        this.selectedProductionView = 'submissions';
        this.filteredRequisitions = this.productionSubmissions.filter(r => r.table_id === table.id);
        this.showToast(`Showing submissions from table: ${table.name}`, 'info');
      }

      this.currentPage = 1;
      this.updatePagination();

    } else if (this.userRole === 'procurement') {
      this.selectedTableId = table.id;
      this.selectedTable = table;
      this.showTableDropdown = false;

      this.filteredRequisitions = this.procurementReviewed.filter(r => r.table_id === table.id);
      // Re-filter summaries for this table
      this.computeProcurementTableSummaries();
      this.currentPage = 1;
      this.updatePagination();

    } else {
      if (table.user_id !== this.userId) {
        this.showToast('You can only access your own tables', 'error');
        return;
      }

      if (table.type !== 'requisition') {
        this.showToast('Invalid table type', 'error');
        return;
      }

      this.selectedTableId = table.id;
      this.selectedTable = table;
      this.showTableDropdown = false;
      this.showAllPending = false;
      this.searchQuery = '';
      this.filterStatus = '';

      localStorage.setItem(`lastSelectedRequisitionTable_${this.userId}`, this.selectedTableId);
      await this.loadRequisitionsDirectly();
    }
  }

  filterByTable(tableId: string) {
    const table = this.tables.find(t => t.id === tableId);
    if (table) {
      this.selectTable(table);
    }
  }

  private async loadTableDetails(tableId: string) {
    try {
      const tableDoc = await this.run(() => getDoc(doc(this.firestore, 'tables', tableId)));
      if (tableDoc.exists()) {
        const data = tableDoc.data();
        this.selectedTable = {
          id: tableDoc.id,
          name: data['name'] || 'Unknown',
          user_id: data['user_id'] || '',
          type: (data['type'] as 'inventory' | 'requisition' | 'production') || 'requisition',
          item_count: data['item_count'] || 0,
          submitted: data['submitted'] || false,
          submitted_at: data['submitted_at'],
          po_file_url: data['po_file_url'],
          po_file_data: data['po_file_data'],
          po_file_mime: data['po_file_mime'],
          po_file_name: data['po_file_name'],
          request_closed: data['request_closed'] || false,
          request_closed_at: data['request_closed_at'],
          request_closed_by: data['request_closed_by']
        };
      }
    } catch (err) {}
  }

  private escapeTableNameRegExp(value: string): string {
    return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  private formatTableDate(date: Date): string {
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const year = date.getFullYear();
    return `${month}/${day}/${year}`;
  }

  private getRequisitionUserLabel(): string {
    const name = (this.userName || '').trim();
    if (!name) return '';
    const atIndex = name.indexOf('@');
    return atIndex > 0 ? name.slice(0, atIndex) : name;
  }

  private getNextRequisitionTableName(): string {
    const dateText = this.formatTableDate(new Date());
    const userLabel = this.getRequisitionUserLabel();
    const prefix = userLabel
      ? `Requisition Slip ${dateText} - ${userLabel}`
      : `Requisition Slip ${dateText}`;
    const escapedPrefix = this.escapeTableNameRegExp(prefix);
    const regex = new RegExp(`^${escapedPrefix}(?: \\((\\d+)\\))?$`);

    let maxSuffix = 0;
    const todayTables = this.tables.filter(table => table.type === 'requisition');
    todayTables.forEach(table => {
      const match = table.name.match(regex);
      if (match) {
        if (match[1]) {
          maxSuffix = Math.max(maxSuffix, Number(match[1]));
        } else {
          maxSuffix = Math.max(maxSuffix, 1);
        }
      }
    });

    return maxSuffix === 0 ? prefix : `${prefix} (${maxSuffix + 1})`;
  }

  openTableModal() {
    this.showTableModal = true;
    if (!this.editingTable) {
      this.newTableName = this.getNextRequisitionTableName();
      this.editTableName = '';
    }
    this.showTableDropdown = false;
  }

  closeTableModal() {
    this.showTableModal = false;
    this.editingTable = null;
  }

  async createTable() {
    if (!this.userId) {
      this.showToast('You must be logged in', 'error');
      return;
    }

    const tableName = this.newTableName.trim() || this.getNextRequisitionTableName();
    this.isSubmitting = true;

    try {
      const result = await this.db.createUserTable({ name: tableName, user_id: this.userId }, 'requisition');

      if (result.success && result.tableId) {
        const newTable: Table = {
          id: result.tableId,
          name: tableName,
          user_id: this.userId,
          type: 'requisition',
          item_count: 0,
          submitted: false,
          created_at: new Date().toISOString()
        };

        this.tables.push(newTable);
        this.selectedTableId = result.tableId;
        this.selectedTable = newTable;

        localStorage.setItem(`lastSelectedRequisitionTable_${this.userId}`, this.selectedTableId);

        this.newTableName = '';
        this.closeTableModal();
        this.showToast(`Table "${tableName}" created successfully`, 'success');
        await this.loadRequisitionsDirectly();
      } else {
        this.showToast('Failed to create table', 'error');
      }
    } catch (err) {
      this.showToast('Failed to create table', 'error');
    } finally {
      this.isSubmitting = false;
    }
  }

  editTable(table: Table) {
    if (table.user_id !== this.userId) {
      this.showToast('You can only edit your own tables', 'error');
      return;
    }
    if (table.type !== 'requisition') {
      this.showToast('Invalid table type', 'error');
      return;
    }
    this.editingTable = table;
    this.editTableName = table.name;
    this.openTableModal();
  }

  async updateTableName() {
    if (!this.editingTable || !this.editTableName.trim()) return;

    try {
      const success = await this.db.updateTableName(
        this.editingTable.id,
        this.editTableName.trim(),
        this.userId
      );

      if (success) {
        const index = this.tables.findIndex(t => t.id === this.editingTable!.id);
        if (index !== -1) this.tables[index].name = this.editTableName.trim();
        if (this.selectedTable?.id === this.editingTable.id) {
          this.selectedTable.name = this.editTableName.trim();
        }
        this.closeTableModal();
        this.showToast('Table renamed successfully', 'success');
      } else {
        this.showToast('Failed to rename table', 'error');
      }
    } catch (err) {
      this.showToast('Failed to rename table', 'error');
    }
  }

  async deleteTable(table: Table) {
    if (this.tables.length <= 1) {
      this.showToast('Cannot delete the last table', 'error');
      return;
    }
    if (table.user_id !== this.userId) {
      this.showToast('You can only delete your own tables', 'error');
      return;
    }
    if (table.type !== 'requisition') {
      this.showToast('Invalid table type', 'error');
      return;
    }
    if (!confirm(`Delete table "${table.name}" and all its requisitions? This cannot be undone.`)) return;

    try {
      const success = await this.db.deleteTable(table.id, this.userId);
      if (success) {
        this.tables = this.tables.filter(t => t.id !== table.id);
        if (this.selectedTableId === table.id) {
          this.selectedTableId = this.tables[0]?.id || '';
          await this.onTableChange();
        }
        this.showToast('Table deleted successfully', 'success');
        if (this.showTableModal) this.closeTableModal();
      } else {
        this.showToast('Failed to delete table', 'error');
      }
    } catch (err) {
      this.showToast('Failed to delete table', 'error');
    }
  }

  openModal() {
    if (!this.selectedTableId && this.viewMode === 'my_tables') {
      this.showToast('Please select a table first', 'error');
      this.openTableModal();
      return;
    }
    this.showModal = true;
    this.submitted = false;
    this.editingRequisition = null;
    this.resetForm();
  }

  openEditModal(req: Requisition) {
    if (!this.selectedTableId && this.viewMode === 'my_tables') {
      this.showToast('Please select a table first', 'error');
      return;
    }

    this.editingRequisition = req;
    this.showModal = true;
    this.submitted = false;
    this.resetForm();

    this.formData = {
      type: req.type || '',
      category: req.category || '',
      skuName: req.skuName || '',
      quantity: req.quantity || null,
      unit: req.unit || '',
      dateNeeded: req.dateNeeded || '',
      supplier: req.supplier || '',
      customSupplier: '',
      brand: req.brand || '',
      customBrand: ''
    };

    const predefinedSuppliers = ['Supplier A', 'Supplier B', 'Supplier C'];
    if (this.formData.supplier && !predefinedSuppliers.includes(this.formData.supplier)) {
      this.formData.customSupplier = this.formData.supplier;
      this.formData.supplier = '__other__';
    }

    const predefinedBrands = ['Brand X', 'Brand Y', 'Brand Z'];
    if (this.formData.brand && !predefinedBrands.includes(this.formData.brand)) {
      this.formData.customBrand = this.formData.brand;
      this.formData.brand = '__other__';
    }

    this.selectedSkuCode = req.skuCode || '';
    if (this.formData.category) this.onCategoryChange();
  }

  async onSubmit() {
    if (!this.selectedTableId && this.viewMode === 'my_tables') {
      this.showToast('Please select a table first', 'error');
      return;
    }
    if (!this.userId) {
      this.showToast('You must be logged in', 'error');
      return;
    }

    this.submitted = true;
    if (!this.validateForm()) {
      this.showToast('Please complete all required fields', 'error');
      return;
    }

    this.isSubmitting = true;

    try {
      const skuName = this.formData.skuName;
      const selectedItem = this.availableSkus.find(item => item.sku_name === skuName);
      const skuCode = this.db.normalizeSkuCode(
        selectedItem ? selectedItem.sku_code : this.selectedSkuCode
      );

      const finalSupplier = this.formData.supplier === '__other__'
        ? (this.formData.customSupplier?.trim() || '')
        : (this.formData.supplier?.trim() || '');

      const finalBrand = this.formData.brand === '__other__'
        ? this.formData.customBrand?.trim()
        : this.formData.brand || '';

      let reqNumber = '';
      if (this.editingRequisition) {
        reqNumber = this.editingRequisition.reqNumber;
      } else {
        const year = new Date().getFullYear();
        const random = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
        reqNumber = `MR-${year}-${random}`;
      }

      const requisitionData: any = {
        reqNumber,
        type: this.formData.type,
        dateNeeded: this.formData.dateNeeded || 'ASAP',
        skuCode,
        skuName,
        quantity: Number(this.formData.quantity),
        unit: this.formData.unit,
        supplier: finalSupplier,
        brand: finalBrand,
        status: this.editingRequisition ? this.editingRequisition.status : 'Pending',
        category: this.formData.category,
        user_id: this.userId,
        table_id: this.selectedTableId || this.editingRequisition?.table_id || '',
        updated_at: new Date().toISOString()
      };

      if (this.editingRequisition && this.editingRequisition.created_at) {
        requisitionData.created_at = this.editingRequisition.created_at;
      }

      let result;
      if (this.editingRequisition) {
        result = await this.db.updateRequisition(
          this.editingRequisition.id,
          requisitionData,
          this.userId,
          this.selectedTableId || this.editingRequisition.table_id || ''
        );
        if (result) this.showToast('Requisition updated successfully', 'success');
      } else {
        result = await this.db.createRequisition(requisitionData, []);
        if (result.success) this.showToast('Requisition created successfully', 'success');
      }

      if (result && (result === true || result.success)) {
        await this.loadRequisitionsDirectly();
        await this.updateTableItemCount();
        this.closeModal();
      } else {
        this.showToast('Failed to save requisition', 'error');
      }
    } catch (err) {
      this.showToast('Failed to save requisition', 'error');
    } finally {
      this.isSubmitting = false;
    }
  }

  async deleteRequisition(req: Requisition) {
    if (!this.selectedTableId || !this.userId) return;

    if (this.userRole !== 'admin' && req.user_id !== this.userId) {
      this.showToast('You can only delete your own requisitions', 'error');
      return;
    }

    if (req.status === 'Approved' || req.status === 'Delivered' || req.status === 'Production_Confirmed') {
      this.showToast('Approved or confirmed requisitions cannot be deleted', 'error');
      return;
    }

    if (!confirm(`Delete requisition ${req.reqNumber || 'Unknown'}?`)) return;

    try {
      const success = await this.db.deleteRequisition(req.id, this.userId, this.selectedTableId);
      if (success) {
        this.requisitions = this.requisitions.filter(r => r.id !== req.id);
        this.applyFilter();
        await this.updateTableItemCount();
        this.showToast('Requisition deleted', 'success');
      } else {
        this.showToast('Could not delete requisition', 'error');
      }
    } catch (err) {
      this.showToast('Delete failed', 'error');
    }
  }

  async submitTable(table: Table) {
    if (this.userRole !== 'user' && this.userRole !== 'store' && this.userRole !== 'admin') {
      this.showToast('Only store/user can submit tables', 'error');
      return;
    }

    if (!confirm(`Submit table "${table.name}" and all its requisitions for approval?`)) return;

    try {
      this.isSubmitting = true;
      this.isLoading = true;

      const snapshot = await this.run(() => {
        const requisitionsRef = collection(this.firestore, 'requisitions');
        const q = query(
          requisitionsRef,
          where('table_id', '==', table.id),
          where('user_id', '==', this.userId)
        );
        return getDocs(q);
      });

      const items: Array<{ skuName: string; skuCode: string; quantity: number; unit: string }> = [];

      await this.run(async () => {
        const batch = writeBatch(this.firestore);
        snapshot.forEach(d => {
          const data = d.data();
          items.push({
            skuName: data['skuName'] || data['sku_name'] || 'Unknown',
            skuCode: data['skuCode'] || data['sku_code'] || 'Unknown',
            quantity: data['quantity'] || data['qty_needed'] || 0,
            unit: data['unit'] || data['batch_unit'] || 'pcs'
          });
          batch.update(d.ref, {
            status: 'Submitted',
            submitted_at: new Date().toISOString()
          });
        });

        const tableRef = doc(this.firestore, 'tables', table.id);
        batch.update(tableRef, {
          submitted: true,
          submitted_at: new Date().toISOString()
        });

        await batch.commit();
      });

      table.submitted = true;
      table.submitted_at = new Date().toISOString();

      const currentUser = await this.auth.getCurrentUserPromise();
      const userEmail = currentUser?.email || 'unknown@example.com';

      try {
        await this.emailNotificationService.sendTableSubmittedNotification({
          tableName: table.name,
          userEmail: userEmail,
          submittedAt: new Date().toISOString(),
          items: items,
          tableId: table.id,
          itemCount: items.length,
          reviewLink: `${window.location.origin}/dashboard/procurement?tableId=${table.id}`
        });
        this.showToast(`Table "${table.name}" submitted and Production team has been notified`, 'success');
      } catch (emailError) {
        this.showToast(`Table "${table.name}" submitted successfully (email notification failed)`, 'info');
      }

      await this.notificationService.sendTableSubmittedNotification(table.id, table.name, this.userId);
      await this.loadRequisitionsDirectly();
    } catch (err) {
      this.showToast('Failed to submit table', 'error');
    } finally {
      this.isSubmitting = false;
      this.isLoading = false;
    }
  }

  // ============================================================
  // Close Request (Procurement only)
  // ============================================================
  async closeRequest(table: Table) {
    if (this.userRole !== 'procurement' && this.userRole !== 'admin') {
      this.showToast('Only procurement can close requests', 'error');
      return;
    }

    if (table.request_closed) {
      this.showToast('This request is already closed', 'info');
      return;
    }

    if (!confirm(`Close request for table "${table.name}"? This will mark it as resolved for all users.`)) return;

    try {
      this.isSubmitting = true;
      const tableRef = doc(this.firestore, 'tables', table.id);
      await this.run(() =>
        updateDoc(tableRef, {
          request_closed: true,
          request_closed_at: new Date().toISOString(),
          request_closed_by: this.userId,
          updated_at: new Date().toISOString()
        })
      );

      // Update local state
      table.request_closed = true;
      table.request_closed_at = new Date().toISOString();
      table.request_closed_by = this.userId;

      const tableIndex = this.tables.findIndex(t => t.id === table.id);
      if (tableIndex !== -1) {
        this.tables[tableIndex] = { ...this.tables[tableIndex], ...table };
      }

      if (this.selectedTable?.id === table.id) {
        this.selectedTable = { ...this.selectedTable, ...table };
      }

      this.showToast(`Request for "${table.name}" has been closed`, 'success');
    } catch (err) {
      this.showToast('Failed to close request', 'error');
    } finally {
      this.isSubmitting = false;
    }
  }

  // ============================================================
  // Raw Materials Modal (User / Production)
  // ============================================================
  async openRawMaterialsModal() {
    if (!this.selectedTable) {
      this.showToast('Please select a table first', 'error');
      return;
    }

    this.showRawMaterialsModal = true;
    this.rawMaterialsModalLoading = true;
    this.rawMaterialsModalData = [];

    try {
      // Determine which requisitions to use
      let reqs: Requisition[] = [];

      if (this.userRole === 'production') {
        // Use productionSubmissions + productionReviewed for the selected table
        reqs = [
          ...this.productionSubmissions.filter(r => r.table_id === this.selectedTable!.id),
          ...this.productionReviewed.filter(r => r.table_id === this.selectedTable!.id)
        ];
        // Deduplicate
        const seen = new Set<string>();
        reqs = reqs.filter(r => {
          if (seen.has(r.id)) return false;
          seen.add(r.id);
          return true;
        });
      } else {
        // user/store — use loaded requisitions
        reqs = this.requisitions;
      }

      // Ensure materials are populated
      await this.populateMaterialsForAllRequisitions(reqs);

      // Load procurement material actions for this table from Firestore
      // We store them in procurementTableSummaries; rebuild the map from it
      await this.loadProcurementActionsForTable(this.selectedTable.id);

      // Aggregate materials across all requisitions
      const materialMap = new Map<string, RawMaterialModalItem>();

      for (const req of reqs) {
        for (const mat of req.materials || []) {
          const key = mat.raw_material.toLowerCase();
          const qty = this.calculateMaterialTotal(req.quantity || req['qty_needed'] || 0, mat.quantity_per_batch);

          const procKey = `${this.selectedTable!.id}|${mat.raw_material.toLowerCase()}`;
          const procAction = this.procurementMaterialActionsMap[procKey] ?? null;

          if (materialMap.has(key)) {
            const existing = materialMap.get(key)!;
            existing.totalQuantity += qty;
            // Production status: if any rejected, mark removed; else use existing
            if (mat.production_action === 'removed') {
              existing.production_status = 'removed';
            } else if (mat.production_action === 'confirmed' && existing.production_status !== 'removed') {
              existing.production_status = 'confirmed';
            }
            // Update procurement action if found
            if (procAction !== null) {
              existing.procurement_action = procAction;
            }
          } else {
            materialMap.set(key, {
              raw_material: mat.raw_material,
              type: mat.type || '—',
              unit: mat.unit || '—',
              totalQuantity: qty,
              production_status: mat.production_action ?? null,
              procurement_action: procAction
            });
          }
        }
      }

      this.rawMaterialsModalData = Array.from(materialMap.values()).sort((a, b) =>
        a.raw_material.localeCompare(b.raw_material)
      );
    } catch (err) {
      this.showToast('Failed to load raw materials', 'error');
    } finally {
      this.rawMaterialsModalLoading = false;
    }
  }

  closeRawMaterialsModal() {
    this.showRawMaterialsModal = false;
    this.rawMaterialsModalData = [];
  }

  /**
   * Load procurement material actions for a given table from the
   * procurementTableSummaries (already in memory) or re-derive from Firestore.
   */
  private async loadProcurementActionsForTable(tableId: string) {
    // Try to get from already-loaded summaries first
    const summary = this.procurementTableSummaries.find(s => s.table_id === tableId);
    if (summary) {
      for (const mat of summary.materials) {
        const key = `${tableId}|${mat.raw_material.toLowerCase()}`;
        this.procurementMaterialActionsMap[key] = mat.procurement_action ?? null;
      }
      return;
    }

    // Otherwise, try to load from Firestore requisitions for this table
    // (procurement actions are stored on the procurementTableSummaries in memory,
    //  but if procurement hasn't been loaded yet, we attempt to load it)
    if (this.procurementReviewed.length === 0) {
      // Minimal load: just get the requisitions for this table to check stored actions
      try {
        const requisitionsRef = collection(this.firestore, 'requisitions');
        const snap = await this.run(() =>
          getDocs(query(requisitionsRef, where('table_id', '==', tableId)))
        );

        snap.forEach(docSnap => {
          const data = docSnap.data();
          const materials: Material[] = data['materials'] || [];
          materials.forEach(mat => {
            if ((mat as any)['procurement_action']) {
              const key = `${tableId}|${(mat.raw_material || '').toLowerCase()}`;
              this.procurementMaterialActionsMap[key] = (mat as any)['procurement_action'];
            }
          });
        });
      } catch (err) {}
    }
  }

  /**
   * Helper used in the procurement exact items modal to get
   * the procurement action for a specific material within a requisition.
   */
  getMaterialProcurementAction(req: Requisition, mat: Material): 'approved' | 'rejected' | null {
    if (!req.table_id) return null;
    const key = `${req.table_id}|${(mat.raw_material || '').toLowerCase()}`;
    return this.procurementMaterialActionsMap[key] ?? null;
  }

  openProductionActionModal(req: Requisition, action: 'confirmed' | 'removed') {
    this.selectedRequisition = req;
    this.productionActionType = action;
    this.productionActionNotes = '';
    this.showProductionActionModal = true;
  }

  async markProductionAction(req: Requisition, action: 'confirmed') {
    try {
      const updateData: any = {
        production_action: action,
        production_action_at: new Date().toISOString(),
        production_action_by: this.userId
      };

      const success = await this.db.updateRequisitionStatus(
        req.id,
        req.status,
        this.userId,
        req.table_id || '',
        updateData
      );

      if (success) {
        req.production_action = action;
        req.production_action_at = updateData.production_action_at;
        req.production_action_by = updateData.production_action_by;
        this.showToast(`Requisition ${action === 'confirmed' ? 'confirmed' : 'marked for removal'}`, 'success');
      } else {
        this.showToast('Failed to update requisition', 'error');
      }
    } catch (err) {
      this.showToast('Failed to update requisition', 'error');
    }
  }

  async confirmProductionAction() {
    if (!this.selectedRequisition) return;

    try {
      const updateData: any = {
        production_action: this.productionActionType,
        production_action_at: new Date().toISOString(),
        production_action_by: this.userId
      };

      if (this.productionActionNotes) {
        updateData.production_action_notes = this.productionActionNotes;
      }

      const success = await this.db.updateRequisitionStatus(
        this.selectedRequisition.id,
        this.selectedRequisition.status,
        this.userId,
        this.selectedRequisition.table_id || '',
        updateData
      );

      if (success) {
        this.selectedRequisition.production_action = this.productionActionType;
        this.selectedRequisition.production_action_at = updateData.production_action_at;
        this.selectedRequisition.production_action_by = updateData.production_action_by;
        if (this.productionActionNotes) {
          this.selectedRequisition.production_action_notes = this.productionActionNotes;
        }
        this.closeProductionActionModal();
        this.showToast(`Requisition marked for ${this.productionActionType === 'confirmed' ? 'confirmation' : 'removal'}`, 'success');
      } else {
        this.showToast('Failed to update requisition', 'error');
      }
    } catch (err) {
      this.showToast('Failed to update requisition', 'error');
    }
  }

  closeProductionActionModal() {
    this.showProductionActionModal = false;
    this.selectedRequisition = null;
    this.productionActionNotes = '';
  }

  async markDelivered(req: Requisition) {
    if (this.userRole !== 'procurement' && this.userRole !== 'admin') {
      this.showToast('Only procurement can mark as delivered', 'error');
      return;
    }
    if (!confirm(`Mark requisition ${req.reqNumber || req.id} as fully delivered?`)) return;

    try {
      const success = await this.db.updateRequisitionStatus(
        req.id,
        'Delivered',
        this.userId,
        req.table_id || '',
        {}
      );
      if (success) {
        req.status = 'Delivered';
        this.showToast('Requisition marked as delivered', 'success');
      } else {
        this.showToast('Failed to update', 'error');
      }
    } catch (err) {
      this.showToast('Failed to update', 'error');
    }
  }

  openMissingNotesModal(req: Requisition) {
    this.selectedRequisition = req;
    this.missingMaterialsNotes = '';
    this.showMissingNotesModal = true;
  }

  async saveMissingNotes() {
    if (!this.selectedRequisition || !this.missingMaterialsNotes.trim()) {
      this.showToast('Please add notes for missing materials', 'error');
      return;
    }
    try {
      const success = await this.db.updateRequisitionStatus(
        this.selectedRequisition.id,
        'Partially_Delivered',
        this.userId,
        this.selectedRequisition.table_id || '',
        { missing_materials_notes: this.missingMaterialsNotes }
      );
      if (success) {
        this.selectedRequisition.status = 'Partially_Delivered';
        this.selectedRequisition.procurement_notes = this.missingMaterialsNotes;
        this.closeMissingNotesModal();
        this.showToast('Notes saved - requisition marked as partially delivered', 'success');
      } else {
        this.showToast('Failed to save notes', 'error');
      }
    } catch (err) {
      this.showToast('Failed to save notes', 'error');
    }
  }

  closeMissingNotesModal() {
    this.showMissingNotesModal = false;
    this.selectedRequisition = null;
    this.missingMaterialsNotes = '';
  }

  openScheduleModal(req: Requisition) {
    if (this.userRole !== 'procurement' && this.userRole !== 'admin') {
      this.showToast('Only procurement can schedule requisitions', 'error');
      return;
    }

    if (req.status !== 'Production_Confirmed') {
      this.showToast('Only confirmed requisitions can be scheduled', 'error');
      return;
    }

    this.selectedRequisition = req;

    if (req.scheduled_date) {
      try {
        const date = new Date(req.scheduled_date);
        this.scheduledDate = date.toISOString().split('T')[0];
        this.scheduledTime = date.toTimeString().split(' ')[0].substring(0, 5);
      } catch (e) {
        this.scheduledDate = '';
        this.scheduledTime = '';
      }
    } else {
      this.scheduledDate = '';
      this.scheduledTime = '';
    }

    this.showScheduleModal = true;
  }

  async scheduleRequisition() {
    if (!this.selectedRequisition || !this.scheduledDate) {
      this.showToast('Please select a date', 'error');
      return;
    }

    try {
      const scheduledDateTime = this.scheduledTime
        ? `${this.scheduledDate}T${this.scheduledTime}`
        : `${this.scheduledDate}T00:00:00`;

      const success = await this.db.updateRequisitionStatus(
        this.selectedRequisition.id,
        'Scheduled',
        this.userId,
        this.selectedRequisition.table_id || this.selectedTableId || '',
        {
          scheduled_date: scheduledDateTime,
          scheduled_by: this.userId
        }
      );

      if (success) {
        await this.loadProcurementReviewed();
        if (this.selectedTable) {
          this.filteredRequisitions = this.procurementReviewed.filter(r => r.table_id === this.selectedTable!.id);
        } else {
          this.filteredRequisitions = [...this.procurementReviewed];
        }
        this.closeScheduleModal();
        this.showToast('Requisition scheduled successfully', 'success');
      } else {
        this.showToast('Failed to schedule requisition', 'error');
      }
    } catch (err) {
      this.showToast('Failed to schedule requisition', 'error');
    }
  }

  closeScheduleModal() {
    this.showScheduleModal = false;
    this.selectedRequisition = null;
    this.scheduledDate = '';
    this.scheduledTime = '';
  }

  openApproveModal(req: Requisition) {
    if (this.userRole !== 'admin') {
      this.showToast('Only admins can approve requisitions', 'error');
      return;
    }
    if (req.status !== 'Scheduled') {
      this.showToast('Only scheduled requisitions can be approved', 'error');
      return;
    }
    this.selectedRequisition = req;
    this.approvalNotes = '';
    this.showApproveModal = true;
  }

  async approveRequisition() {
    if (!this.selectedRequisition) return;

    try {
      const success = await this.db.updateRequisitionStatus(
        this.selectedRequisition.id,
        'Approved',
        this.userId,
        this.selectedTableId || '',
        {
          approved_by: this.userId,
          approval_notes: this.approvalNotes || null,
          approved_at: new Date().toISOString()
        }
      );

      if (success) {
        await this.loadRequisitionsDirectly();
        this.closeApproveModal();
        this.showToast('Requisition approved successfully', 'success');
      } else {
        this.showToast('Failed to approve requisition', 'error');
      }
    } catch (err) {
      this.showToast('Failed to approve requisition', 'error');
    }
  }

  closeApproveModal() {
    this.showApproveModal = false;
    this.selectedRequisition = null;
    this.approvalNotes = '';
  }

  async submitReviewedTable() {
    if (!this.canSubmitReviewedTable() || !this.selectedTable) return;
    if (!confirm(`Transfer reviewed table "${this.selectedTable.name}" to procurement?`)) return;

    try {
      this.isSubmitting = true;
      this.isLoading = true;

      const tableSubmissions = this.productionSubmissions.filter(r => r.table_id === this.selectedTable!.id);
      const confirmedItems = tableSubmissions.filter(r => r.production_action === 'confirmed').length;
      const removedItems = tableSubmissions.filter(r => r.production_action === 'removed').length;

      const updatePromises = tableSubmissions.map(async (req) => {
        const action = req.production_action === 'removed' ? 'removed' : 'confirmed';
        const newStatus = action === 'removed' ? 'Removed' : 'Production_Confirmed';

        const updateData: any = {
          production_action: action,
          production_action_at: req.production_action_at || new Date().toISOString(),
          production_action_by: req.production_action_by || this.userId
        };

        if (action === 'removed') {
          updateData.production_action_notes = req.production_action_notes || '';
        }

        const needsUpdate = req.status !== newStatus || req.production_action !== action;
        if (!needsUpdate) return true;

        const tableId = req.table_id || this.selectedTableId || '';
        return this.db.updateRequisitionStatus(req.id, newStatus, this.userId, tableId, updateData);
      });

      await Promise.all(updatePromises);

      const tableRef = doc(this.firestore, 'tables', this.selectedTable.id);
      await this.run(() =>
        updateDoc(tableRef, {
          production_reviewed: true,
          production_reviewed_at: new Date().toISOString(),
          production_reviewed_by: this.userId,
          updated_at: new Date().toISOString(),
          submitted: true
        })
      );

      const currentUser = await this.auth.getCurrentUserPromise();
      const reviewerEmail = currentUser?.email || 'unknown@example.com';

      try {
        await this.emailNotificationService.sendTableReviewedNotification({
          tableName: this.selectedTable.name,
          reviewerEmail: reviewerEmail,
          reviewedAt: new Date().toISOString(),
          totalItems: tableSubmissions.length,
          confirmedItems: confirmedItems,
          removedItems: removedItems,
          tableId: this.selectedTable.id,
          reviewLink: `${window.location.origin}/requisitions?tableId=${this.selectedTable.id}&role=procurement`
        });
        this.showToast(`Table "${this.selectedTable.name}" transferred to procurement – they have been notified`, 'success');
      } catch (emailError) {
        this.showToast(`Table "${this.selectedTable.name}" transferred to procurement (email notification failed)`, 'info');
      }

      await this.notificationService.sendTableReviewedByProductionNotification(
        this.selectedTable.id,
        this.selectedTable.name,
        this.userId
      );

      if (this.selectedTable) {
        this.selectedTable.production_reviewed = true;
      }

      await this.loadProductionSubmissions();
      await this.loadProductionReviewed();

      this.selectedProductionView = 'reviewed';
      this.filteredRequisitions = this.productionReviewed.filter(r => r.table_id === this.selectedTable!.id);
      this.updatePagination();
    } catch (err) {
      this.showToast('Failed to submit table', 'error');
    } finally {
      this.isSubmitting = false;
      this.isLoading = false;
    }
  }

  async toggleRow(req: Requisition) {
    if (!req.id) return;
    this.expandedRows[req.id] = !this.expandedRows[req.id];
    this.cdr.detectChanges();
  }

  async confirmAllMaterials(req: Requisition) {
    if (!req || !req.reqNumber) return;
    if (!confirm(`Accept all materials for requisition ${req.reqNumber}?`)) return;
    await this.setAllMaterialsAction(req, 'confirmed');
  }

  async rejectAllMaterials(req: Requisition) {
    if (!req || !req.reqNumber) return;
    if (!confirm(`Reject all materials for requisition ${req.reqNumber}?`)) return;
    await this.setAllMaterialsAction(req, 'removed');
  }

  private async setAllMaterialsAction(req: Requisition, action: 'confirmed' | 'removed') {
    if (!req) return;

    if (!req.materials || req.materials.length === 0) {
      const skuCode = String(req.skuCode ?? req['sku_code'] ?? '').trim();
      const skuName = String(req.skuName ?? req['sku_name'] ?? '').trim();
      req.materials = this.getMaterialsFromLoadedData(skuCode, skuName);
    }

    if (!req.materials || req.materials.length === 0) {
      this.showToast('No materials were found for this requisition.', 'error');
      return;
    }

    req.materials = req.materials.map(mat => ({ ...mat, production_action: action }));
    req.materials = [...req.materials];

    req.production_action = action;
    req.production_action_at = new Date().toISOString();
    req.production_action_by = this.userId;

    const updateData: any = {
      materials: req.materials,
      production_action: req.production_action,
      production_action_at: req.production_action_at,
      production_action_by: req.production_action_by
    };

    const success = await this.db.updateRequisitionStatus(
      req.id,
      req.status,
      this.userId,
      req.table_id || '',
      updateData
    );

    if (success) {
      this.showToast(`All materials ${action === 'confirmed' ? 'accepted' : 'rejected'} for requisition ${req.reqNumber}`, 'success');
    } else {
      this.showToast(`Failed to ${action === 'confirmed' ? 'accept' : 'reject'} materials`, 'error');
    }
  }

  async setMaterialAction(req: Requisition, material: Material, action: 'confirmed' | 'removed') {
    if (!req || !req.materials) return;

    material.production_action = action;
    req.materials = [...req.materials];

    const allMaterialsCompleted = req.materials.every(
      mat => mat.production_action === 'confirmed' || mat.production_action === 'removed'
    );
    const rejectedCount = req.materials.filter(mat => mat.production_action === 'removed').length;

    if (allMaterialsCompleted) {
      req.production_action = rejectedCount === req.materials.length ? 'removed' : 'confirmed';
      req.production_action_at = new Date().toISOString();
      req.production_action_by = this.userId;
    } else {
      delete req.production_action;
      delete req.production_action_at;
      delete req.production_action_by;
    }

    const updateData: any = { materials: req.materials };
    if (req.production_action) {
      updateData.production_action = req.production_action;
      updateData.production_action_at = req.production_action_at;
      updateData.production_action_by = req.production_action_by;
    }

    const success = await this.db.updateRequisitionStatus(
      req.id,
      req.status,
      this.userId,
      req.table_id || '',
      updateData
    );

    if (success) {
      this.showToast(
        `Material ${material.raw_material} ${action === 'confirmed' ? 'accepted' : 'rejected'} for requisition ${req.reqNumber}`,
        'success'
      );
    } else {
      this.showToast(`Failed to update material action`, 'error');
    }
  }

  private computeProcurementTableSummaries() {
    const tableMap = new Map<string, ProcurementTableSummary>();

    // Filter requisitions to the currently selected table if one is selected
    const reqs = this.selectedTableId
      ? this.procurementReviewed.filter(r => r.table_id === this.selectedTableId)
      : this.procurementReviewed;

    for (const req of reqs) {
      if (!req.table_id) continue;
      if (req.status === 'Removed' || req.production_action === 'removed') continue;

      const tableName = req.table_name || this.tableNameMap[req.table_id] || 'Unknown Table';

      if (!req.materials || req.materials.length === 0) {
        const skuCode = String(req.skuCode ?? req['sku_code'] ?? '').trim();
        const skuName = String(req.skuName ?? req['sku_name'] ?? '').trim();
        req.materials = this.getMaterialsFromLoadedData(skuCode, skuName);
      }

      let summary = tableMap.get(req.table_id);
      if (!summary) {
        summary = {
          table_id: req.table_id,
          table_name: tableName,
          uniqueMaterialsCount: 0,
          totalRequestedQuantity: 0,
          materials: []
        };
        tableMap.set(req.table_id, summary);
      }

      for (const mat of req.materials || []) {
        if (mat.production_action === 'removed') continue;

        const qty = this.calculateMaterialTotal(req.quantity || req['qty_needed'] || 0, mat.quantity_per_batch);
        const matKey = mat.raw_material.toLowerCase();
        const procKey = `${req.table_id}|${matKey}`;
        const existingProcAction = this.procurementMaterialActionsMap[procKey] ?? null;

        let existing = summary.materials.find(
          item =>
            item.raw_material.toLowerCase() === matKey &&
            item.unit === mat.unit &&
            item.type === mat.type
        );

        if (!existing) {
          existing = {
            raw_material: mat.raw_material,
            unit: mat.unit || '—',
            type: mat.type || 'N/A',
            totalQuantity: qty,
            table_id: req.table_id,
            table_name: tableName,
            procurement_action: existingProcAction,
            production_status: mat.production_action ?? null
          };
          summary.materials.push(existing);
        } else {
          existing.totalQuantity += qty;
          if ((mat.production_action as string) === 'removed') {
            existing.production_status = 'removed';
          } else if (mat.production_action === 'confirmed' && existing.production_status !== 'removed') {
            existing.production_status = 'confirmed';
          }
          if (existingProcAction !== null) {
            existing.procurement_action = existingProcAction;
          }
        }
      }
    }

    this.procurementTableSummaries = Array.from(tableMap.values()).map(summary => ({
      ...summary,
      uniqueMaterialsCount: summary.materials.length,
      totalRequestedQuantity: summary.materials.reduce((sum, mat) => sum + (mat.totalQuantity || 0), 0)
    }));
  }

  setProcurementMaterialAction(item: ProcurementMaterialSummary, action: 'approved' | 'rejected') {
    item.procurement_action = action;
    // Persist to the actions map so other views can read it
    const key = `${item.table_id}|${item.raw_material.toLowerCase()}`;
    this.procurementMaterialActionsMap[key] = action;

    // Optionally save to Firestore on each click for realtime visibility
    this.saveProcurementMaterialActionToFirestore(item.table_id, item.raw_material, action);
  }

  private async saveProcurementMaterialActionToFirestore(
    tableId: string,
    rawMaterial: string,
    action: 'approved' | 'rejected'
  ) {
    try {
      // Find all procurement requisitions for this table that contain this raw material
      const tableReqs = this.procurementReviewed.filter(r => r.table_id === tableId);
      for (const req of tableReqs) {
        const mat = req.materials?.find(
          m => m.raw_material.toLowerCase() === rawMaterial.toLowerCase()
        );
        if (mat) {
          // Store the procurement action on the material in Firestore
          const updatedMaterials = (req.materials || []).map(m => {
            if (m.raw_material.toLowerCase() === rawMaterial.toLowerCase()) {
              return { ...m, procurement_action: action };
            }
            return m;
          });

          const reqRef = doc(this.firestore, 'requisitions', req.id);
          await this.run(() => updateDoc(reqRef, { materials: updatedMaterials }));
          req.materials = updatedMaterials;
        }
      }
    } catch (err) {
      // Non-critical: actions already saved in memory
      console.error('Failed to persist procurement action to Firestore', err);
    }
  }

  openProcurementOriginalItemsModal(summary: ProcurementTableSummary) {
    this.currentProcurementTableId = summary.table_id;
    this.currentProcurementTableName = summary.table_name;
    this.expandedProcurementModalRows = {};
    this.showProcurementOriginalItemsModal = true;
  }

  closeProcurementOriginalItemsModal() {
    this.showProcurementOriginalItemsModal = false;
    this.currentProcurementTableId = '';
    this.currentProcurementTableName = '';
    this.expandedProcurementModalRows = {};
  }

  toggleProcurementModalRow(reqId: string) {
    this.expandedProcurementModalRows[reqId] = !this.expandedProcurementModalRows[reqId];
  }

  getProcurementOriginalRequisitions(tableId: string) {
    return this.procurementReviewed.filter(req => req.table_id === tableId);
  }

  getTableForSummary(summary: ProcurementTableSummary): Table | undefined {
    return this.tables.find(t => t.id === summary.table_id);
  }

  hasPoDocument(table: Table | null | undefined): boolean {
    if (!table) return false;
    return !!(table.po_file_data || table.po_file_url);
  }

  getPoViewUrl(table: Table | null | undefined): string {
    if (!table) return '';
    if (table.po_file_data) {
      const mime = table.po_file_mime || table.po_file_type || 'application/octet-stream';
      return `data:${mime};base64,${table.po_file_data}`;
    }
    return table.po_file_url || '';
  }

  private getMaterialsFromLoadedData(skuCode: string, skuName: string): Material[] {
    if (!this.masterDataRows || this.masterDataRows.length === 0) return [];

    const skuCodeLower = skuCode.toLowerCase().trim();
    const skuNameLower = skuName.toLowerCase().trim();

    const materials: Material[] = [];
    const seen = new Set<string>();

    this.masterDataRows.forEach(row => {
      const rowSkuCode = (row.sku_code || '').toString().trim().toLowerCase();
      const rowSkuName = (row.sku_name || '').toString().trim().toLowerCase();

      if (
        (skuCodeLower && rowSkuCode === skuCodeLower) ||
        (skuNameLower && rowSkuName === skuNameLower)
      ) {
        const rawMaterial = (row.raw_material || '').trim();
        if (rawMaterial && !seen.has(rawMaterial.toLowerCase())) {
          materials.push({
            raw_material: rawMaterial,
            quantity_per_batch: row.qty_per_batch ?? null,
            unit: (row.batch_unit || '').trim(),
            type: (row.type || '').trim()
          });
          seen.add(rawMaterial.toLowerCase());
        }
      }
    });

    return materials;
  }

  calculateMaterialTotal(quantity: number, qtyPerBatch: number | null): number {
    const qty = quantity || 0;
    const batchQty = qtyPerBatch || 0;
    return batchQty * qty;
  }

  async onCategoryChange() {
    if (!this.formData.category) {
      this.availableSkus = [];
      this.formData.skuName = '';
      this.selectedSkuCode = '';
      return;
    }

    try {
      this.availableSkus = await this.db.getSkusByCategory(this.formData.category);
      if (this.editingRequisition && this.formData.skuName) {
        const skuExists = this.availableSkus.some(s => s.sku_name === this.formData.skuName);
        if (!skuExists) {
          this.formData.skuName = '';
          this.selectedSkuCode = '';
        } else {
          const selectedItem = this.availableSkus.find(s => s.sku_name === this.formData.skuName);
          this.selectedSkuCode = selectedItem ? selectedItem.sku_code : '';
        }
      } else {
        this.formData.skuName = '';
        this.selectedSkuCode = '';
      }
    } catch (err) {
      this.showToast('Could not load SKUs', 'error');
    }
  }

  onSkuNameSelect() {
    if (!this.formData.skuName) {
      this.selectedSkuCode = '';
      return;
    }
    const selectedItem = this.availableSkus.find(item => item.sku_name === this.formData.skuName);
    this.selectedSkuCode = selectedItem ? selectedItem.sku_code : '';
  }

  // P.O File Upload Methods
  triggerPoFileInput(table?: Table | null) {
    this.poUploadTargetTable = table || this.selectedTable;
    const fileInput = document.getElementById('poFileInput') as HTMLInputElement;
    if (fileInput) {
      fileInput.value = '';
      fileInput.click();
    }
  }

  onPoFileSelected(event: Event) {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (file) {
      if (file.size > this.maxPoFileBytes) {
        this.showToast(`File must be less than ${Math.round(this.maxPoFileBytes / 1024)}KB (Firestore limit)`, 'error');
        this.poFile = null;
        this.poFileName = '';
        return;
      }

      const allowedTypes = ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.jpg', '.jpeg', '.png'];
      const fileExt = '.' + file.name.split('.').pop()?.toLowerCase();
      if (!allowedTypes.includes(fileExt)) {
        this.showToast('Please upload PDF, Word, Excel, or image files only', 'error');
        this.poFile = null;
        this.poFileName = '';
        return;
      }

      this.poFile = file;
      this.poFileName = file.name;
    } else {
      this.poFile = null;
      this.poFileName = '';
    }
  }

  private readFileAsBase64(file: File): Promise<{ base64: string; mime: string }> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result as string;
        const match = result.match(/^data:([^;]+);base64,(.+)$/);
        if (match) {
          resolve({ mime: match[1], base64: match[2] });
        } else {
          reject(new Error('Failed to read file'));
        }
      };
      reader.onerror = () => reject(new Error('Failed to read file'));
      reader.readAsDataURL(file);
    });
  }

  async uploadPoFile(table?: Table | null) {
    const targetTable = table || this.poUploadTargetTable || this.selectedTable;
    if (!targetTable) {
      this.showToast('Please select a table first', 'error');
      return;
    }
    if (!this.poFile) {
      this.showToast('Please choose a P.O file first', 'error');
      return;
    }

    try {
      this.isUploadingPo = true;
      this.showToast('Saving P.O to Firestore...', 'info');

      const { base64, mime } = await this.readFileAsBase64(this.poFile);
      const tableRef = doc(this.firestore, 'tables', targetTable.id);
      await this.run(() =>
        updateDoc(tableRef, {
          po_file_data: base64,
          po_file_mime: mime,
          po_file_name: this.poFile?.name || '',
          po_file_size: this.poFile?.size || 0,
          po_file_type: mime,
          po_file_url: null,
          po_uploaded_at: new Date().toISOString(),
          po_uploaded_by: this.userId,
          updated_at: new Date().toISOString()
        })
      );

      targetTable.po_file_data = base64;
      targetTable.po_file_mime = mime;
      targetTable.po_file_name = this.poFile?.name || '';
      targetTable.po_file_size = this.poFile?.size;
      targetTable.po_file_type = mime;
      targetTable.po_file_url = undefined;

      const tableIndex = this.tables.findIndex(t => t.id === targetTable.id);
      if (tableIndex !== -1) {
        this.tables[tableIndex] = { ...this.tables[tableIndex], ...targetTable };
      }
      if (this.selectedTable?.id === targetTable.id) {
        this.selectedTable = { ...this.selectedTable, ...targetTable };
      }

      this.poFile = null;
      this.poFileName = '';
      this.poUploadTargetTable = null;
      const fileInput = document.getElementById('poFileInput') as HTMLInputElement;
      if (fileInput) fileInput.value = '';

      this.showToast('P.O uploaded and saved successfully', 'success');
    } catch (err: any) {
      this.showToast(err?.message || 'Failed to save P.O file. Please try again.', 'error');
    } finally {
      this.isUploadingPo = false;
    }
  }

  async removePoLink(table?: Table | null) {
    const targetTable = table || this.selectedTable;
    if (!targetTable) return;
    if (!confirm('Remove this P.O document?')) return;

    try {
      this.isSubmitting = true;
      this.showToast('Removing P.O document...', 'info');

      const tableRef = doc(this.firestore, 'tables', targetTable.id);
      await this.run(() =>
        updateDoc(tableRef, {
          po_file_data: null,
          po_file_mime: null,
          po_file_url: null,
          po_file_name: null,
          po_file_size: null,
          po_file_type: null,
          po_removed_at: new Date().toISOString(),
          po_removed_by: this.userId,
          updated_at: new Date().toISOString()
        })
      );

      targetTable.po_file_data = undefined;
      targetTable.po_file_mime = undefined;
      targetTable.po_file_url = undefined;
      targetTable.po_file_name = undefined;
      targetTable.po_file_size = undefined;
      targetTable.po_file_type = undefined;

      const tableIndex = this.tables.findIndex(t => t.id === targetTable.id);
      if (tableIndex !== -1) {
        this.tables[tableIndex] = { ...this.tables[tableIndex], ...targetTable };
      }
      if (this.selectedTable?.id === targetTable.id) {
        this.selectedTable = { ...this.selectedTable, ...targetTable };
      }

      this.showToast('P.O document removed successfully', 'success');
    } catch (err) {
      this.showToast('Failed to remove P.O document', 'error');
    } finally {
      this.isSubmitting = false;
    }
  }

  canUploadMasterData(): boolean {
    return this.userRole === 'user' || this.userRole === 'store' || this.userRole === 'admin';
  }

  canViewMasterData(): boolean {
    return (
      this.userRole === 'user' ||
      this.userRole === 'store' ||
      this.userRole === 'admin' ||
      this.userRole === 'production' ||
      this.userRole === 'procurement'
    );
  }

  async openMasterDataModal() {
    this.showMasterDataModal = true;
    this.masterDataSearchQuery = '';
    this.loadingMasterDataView = true;
    this.masterDataRows = [];
    this.filteredMasterDataRows = [];
    this.groupedMasterData = {};
    this.groupedMasterDataArray = [];
    this.filteredGroupedMasterData = [];
    this.expandedSkus = {};

    try {
      this.masterDataRows = await this.db.getAllMasterData();
      this.buildGroupedMasterData();
      this.applyMasterDataFilter();
    } catch {
      this.showToast('Failed to load master data', 'error');
    } finally {
      this.loadingMasterDataView = false;
    }
  }

  private buildGroupedMasterData() {
    this.groupedMasterData = {};
    this.masterDataRows.forEach(row => {
      const skuCode = (row.sku_code || '').trim();
      const skuName = (row.sku_name || '').trim();
      const skuKey = `${skuCode}|${skuName}`;

      if (!this.groupedMasterData[skuKey]) {
        this.groupedMasterData[skuKey] = { sku: { code: skuCode, name: skuName }, materials: [] };
      }
      this.groupedMasterData[skuKey].materials.push(row);
    });

    this.groupedMasterDataArray = Object.entries(this.groupedMasterData).map(([skuKey, data]) => ({
      skuKey,
      sku: data.sku,
      materials: data.materials,
      materialCount: data.materials.length
    })).sort((a, b) => {
      const catA = (a.materials[0]?.category || '').localeCompare(b.materials[0]?.category || '');
      if (catA !== 0) return catA;
      return a.sku.code.localeCompare(b.sku.code);
    });
  }

  closeMasterDataModal() {
    this.showMasterDataModal = false;
    this.masterDataSearchQuery = '';
    this.expandedSkus = {};
  }

  toggleSkuExpand(skuKey: string) {
    this.expandedSkus[skuKey] = !this.expandedSkus[skuKey];
    this.cdr.detectChanges();
  }

  applyMasterDataFilter() {
    const q = this.masterDataSearchQuery.trim().toLowerCase();

    if (!q) {
      this.filteredMasterDataRows = [...this.masterDataRows];
      this.filteredGroupedMasterData = [...this.groupedMasterDataArray];
      return;
    }

    this.filteredGroupedMasterData = this.groupedMasterDataArray.filter(group => {
      if (
        (group.sku.code || '').toLowerCase().includes(q) ||
        (group.sku.name || '').toLowerCase().includes(q)
      ) {
        return true;
      }
      return group.materials.some(mat =>
        (mat.category || '').toLowerCase().includes(q) ||
        (mat.raw_material || '').toLowerCase().includes(q) ||
        (mat.supplier || '').toLowerCase().includes(q)
      );
    });

    this.filteredMasterDataRows = this.masterDataRows.filter(row =>
      (row.sku_code || '').toLowerCase().includes(q) ||
      (row.sku_name || '').toLowerCase().includes(q) ||
      (row.category || '').toLowerCase().includes(q) ||
      (row.raw_material || '').toLowerCase().includes(q) ||
      (row.supplier || '').toLowerCase().includes(q)
    );
  }

  exportMasterDataToXlsx() {
    const rows = this.filteredMasterDataRows.length ? this.filteredMasterDataRows : this.masterDataRows;
    if (!rows.length) {
      this.showToast('No master data to export', 'error');
      return;
    }

    const headers = [
      'SKU Code', 'SKU Name', 'Qty Per Unit', 'Unit', 'Qty Per Pack', 'Pack Unit',
      'Projected Yield Per Batch', 'Yield Unit', 'Category', 'Raw Material',
      'Qty Per Batch', 'Batch Unit', 'Type', 'Supplier'
    ];

    const data = rows.map(row => [
      row.sku_code || '', row.sku_name || '', row.qty_per_unit ?? '', row.unit || '',
      row.qty_per_pack ?? '', row.pack_unit || '', row.projected_yield_per_batch ?? '',
      row.yield_unit || '', row.category || '', row.raw_material || '',
      row.qty_per_batch ?? '', row.batch_unit || '', row.type || '', row.supplier || ''
    ]);

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...data]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Master Data');

    const date = new Date().toISOString().split('T')[0];
    XLSX.writeFile(workbook, `master-data-${date}.xlsx`);
    this.showToast('Master data exported', 'success');
  }

  async onFileSelected(event: any) {
    const file = event.target.files?.[0];
    if (!file) return;

    this.selectedFileName = file.name;
    this.importStatus = 'loading';
    this.importMessage = 'Uploading...';

    try {
      const result = await this.db.uploadMasterData(file);
      if (result.success) {
        this.importStatus = 'success';
        this.importMessage = `Imported ${result.count} rows`;
        await this.loadCategories();
        if (this.showMasterDataModal) {
          this.masterDataRows = await this.db.getAllMasterData();
          this.applyMasterDataFilter();
        }
        this.showToast('Master data imported successfully', 'success');
      } else {
        this.importStatus = 'error';
        this.importMessage = result.error || 'Upload failed';
        this.showToast(result.error || 'Upload failed', 'error');
      }
    } catch (err) {
      this.importStatus = 'error';
      this.importMessage = 'Upload error';
      this.showToast('Upload error', 'error');
    }
  }

  switchProductionView(view: 'submissions' | 'reviewed') {
    this.selectedProductionView = view;
    this.selectedTableId = '';
    this.selectedTable = null;

    if (view === 'submissions') {
      this.filteredRequisitions = [...this.productionSubmissions];
    } else {
      this.filteredRequisitions = [...this.productionReviewed];
    }

    this.currentPage = 1;
    this.updatePagination();
  }

  async updateTableItemCount() {
    if (!this.selectedTableId || !this.userId) return;

    try {
      await this.db.updateTableItemCount(
        this.selectedTableId,
        this.requisitions.length,
        this.userId
      );

      if (this.selectedTable) {
        this.selectedTable.item_count = this.requisitions.length;
      }

      const tableIndex = this.tables.findIndex(t => t.id === this.selectedTableId);
      if (tableIndex !== -1) {
        this.tables[tableIndex].item_count = this.requisitions.length;
      }
    } catch (err) {}
  }

  validateForm(): boolean {
    return !!(
      this.formData.type &&
      this.formData.category &&
      this.formData.skuName &&
      this.formData.quantity &&
      this.formData.quantity > 0 &&
      this.formData.unit
    );
  }

  closeModal() {
    this.showModal = false;
    this.editingRequisition = null;
  }

  resetForm() {
    this.formData = {
      type: '',
      category: '',
      skuName: '',
      quantity: null,
      unit: '',
      dateNeeded: '',
      supplier: '',
      customSupplier: '',
      brand: '',
      customBrand: ''
    };
    this.selectedSkuCode = '';
  }

  onSupplierChange() {
    if (this.formData.supplier !== '__other__') this.formData.customSupplier = '';
  }

  onBrandChange() {
    if (this.formData.brand !== '__other__') this.formData.customBrand = '';
  }

  toggleTableDropdown() {
    this.showTableDropdown = !this.showTableDropdown;
  }

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const target = event.target as HTMLElement;
    if (!target.closest('.dropdown')) {
      this.showTableDropdown = false;
    }
  }

  canManageTables(): boolean {
    return (
      (this.userRole === 'user' || this.userRole === 'store' || this.userRole === 'admin') &&
      this.viewMode === 'my_tables'
    );
  }

  canCreateRequisition(): boolean {
    return (
      this.canManageTables() &&
      !!this.selectedTable &&
      !this.selectedTable.submitted
    );
  }

  canSubmitTable(): boolean {
    return (
      (this.userRole === 'user' || this.userRole === 'store' || this.userRole === 'admin') &&
      this.viewMode === 'my_tables' &&
      this.selectedTable !== null &&
      !this.selectedTable.submitted &&
      this.requisitions.length > 0
    );
  }

  canSubmitReviewedTable(): boolean {
    if (this.userRole !== 'production' || !this.selectedTable) return false;
    if (this.selectedTable.production_reviewed) return false;

    const tableSubmissions = this.productionSubmissions.filter(r => r.table_id === this.selectedTable!.id);
    if (tableSubmissions.length === 0) return false;

    return tableSubmissions.every(
      r => r.production_action === 'confirmed' || r.production_action === 'removed'
    );
  }

  canSubmitRequisition(req: Requisition): boolean {
    return (
      (this.userRole === 'user' || this.userRole === 'store' || this.userRole === 'admin') &&
      req.status === 'Pending' &&
      req.user_id === this.userId &&
      this.selectedTable !== null &&
      !this.selectedTable.submitted
    );
  }

  canProductionAction(req: Requisition): boolean {
    return this.userRole === 'production' && req.status === 'Submitted';
  }

  canMarkDelivered(req: Requisition): boolean {
    return (
      (this.userRole === 'procurement' || this.userRole === 'admin') &&
      (req.status === 'Production_Confirmed' || req.status === 'Partially_Delivered')
    );
  }

  canAddMissingNotes(req: Requisition): boolean {
    return (
      (this.userRole === 'procurement' || this.userRole === 'admin') &&
      req.status === 'Production_Confirmed'
    );
  }

  canScheduleRequisition(req: Requisition): boolean {
    return (
      (this.userRole === 'procurement' || this.userRole === 'admin') &&
      req.status === 'Production_Confirmed'
    );
  }

  canApproveRequisition(req: Requisition): boolean {
    return this.userRole === 'admin' && req.status === 'Scheduled';
  }

  canEditRequisition(req: Requisition): boolean {
    return (
      this.viewMode === 'my_tables' &&
      (this.userRole === 'admin' || req.user_id === this.userId) &&
      req.status !== 'Submitted' &&
      req.status !== 'Approved' &&
      req.status !== 'Rejected' &&
      req.status !== 'Delivered' &&
      req.status !== 'Partially_Delivered' &&
      this.selectedTable !== null &&
      !this.selectedTable.submitted
    );
  }

  canDeleteRequisition(req: Requisition): boolean {
    return (
      this.viewMode === 'my_tables' &&
      (
        this.userRole === 'admin' ||
        ((this.userRole === 'user' || this.userRole === 'store') && req.user_id === this.userId)
      ) &&
      req.status !== 'Submitted' &&
      req.status !== 'Approved' &&
      req.status !== 'Delivered' &&
      this.selectedTable !== null &&
      !this.selectedTable.submitted
    );
  }

  getStatusBadgeClass(status: string): string {
    switch (status) {
      case 'Pending': return 'status-pending';
      case 'Submitted': return 'status-submitted';
      case 'Scheduled': return 'status-scheduled';
      case 'Approved': return 'status-approved';
      case 'Rejected': return 'status-rejected';
      case 'Production_Confirmed': return 'status-scheduled';
      case 'Removed': return 'status-rejected';
      case 'Delivered': return 'status-approved';
      case 'Partially_Delivered': return 'status-pending';
      default: return 'status-pending';
    }
  }

  applyFilter() {
    let filtered = [...this.requisitions];

    if (this.searchQuery.trim()) {
      const q = this.searchQuery.toLowerCase();
      filtered = filtered.filter(r =>
        (r.reqNumber || '').toLowerCase().includes(q) ||
        (r.skuCode || r['sku_code'] || '').toLowerCase().includes(q) ||
        (r.skuName || r['sku_name'] || '').toLowerCase().includes(q) ||
        (r.supplier || '').toLowerCase().includes(q)
      );
    }

    if (this.filterStatus) {
      filtered = filtered.filter(r => r.status === this.filterStatus);
    }

    this.filteredRequisitions = filtered;
    this.currentPage = 1;
    this.updatePagination();
  }

  updatePagination() {
    this.totalPages = Math.max(1, Math.ceil(this.filteredRequisitions.length / this.pageSize));
    const start = (this.currentPage - 1) * this.pageSize;
    this.paginatedRequisitions = this.filteredRequisitions.slice(start, start + this.pageSize);
  }

  goToPage(page: number) {
    if (page >= 1 && page <= this.totalPages) {
      this.currentPage = page;
      this.updatePagination();
    }
  }

  onPageSizeChange() {
    this.currentPage = 1;
    this.updatePagination();
  }

  showToast(message: string, type: 'success' | 'error' | 'info' = 'info') {
    this.snackbarMessage = message;
    this.snackbarType = type;
    this.showSnackbar = true;

    if (this.snackbarTimeout) clearTimeout(this.snackbarTimeout);

    this.snackbarTimeout = setTimeout(() => {
      this.hideSnackbar();
    }, 3000);
  }

  hideSnackbar() {
    this.showSnackbar = false;
    if (this.snackbarTimeout) {
      clearTimeout(this.snackbarTimeout);
      this.snackbarTimeout = null;
    }
  }
}