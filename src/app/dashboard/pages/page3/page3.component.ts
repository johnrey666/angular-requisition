import { Component, OnInit, OnDestroy, HostListener, Injector, ChangeDetectorRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClientModule } from '@angular/common/http';
import { runInInjectionContext } from '@angular/core';
import { DatabaseService } from '../../../core/services/database.service';
import { AuthService } from '../../../core/services/auth.service';
import { NotificationService } from '../../../core/services/notification.service';
import { EmailNotificationService } from '../../../core/services/email-notification.service';
import { ToastService } from '../../../core/services/toast.service';
import {
  Firestore, doc, collection, query, where, getDocs,
  writeBatch, getDoc, updateDoc, orderBy, addDoc, onSnapshot, Unsubscribe, limit
} from '@angular/fire/firestore';
import { Router, ActivatedRoute } from '@angular/router';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';
import * as XLSX from 'xlsx';

interface Material {
  raw_material: string;
  quantity_per_batch: number | null;
  unit: string;
  type: string;
  production_action?: 'confirmed' | 'removed';
  procurement_action?: 'approved' | 'rejected';
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
  production_notes?: string;
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
  sm_file_url?: string;
  sm_file_data?: string;
  sm_file_mime?: string;
  sm_file_name?: string;
  sm_file_size?: number;
  sm_file_type?: string;
  sm_uploaded_at?: string;
  sm_uploaded_by?: string;
  production_reviewed?: boolean;
  production_reviewed_at?: string;
  production_reviewed_by?: string;
  procurement_date_needed?: string;
  request_closed?: boolean;
  request_closed_at?: string;
  request_closed_by?: string;
  request_status?: 'PENDING' | 'DONE';
  close_request_pending?: boolean;
  close_request_user_accepted?: boolean;
  close_request_production_accepted?: boolean;
  close_request_initiated_at?: string;
  close_request_initiated_by?: string;
  close_request_mode?: 'procurement' | 'production';
  procurement_batch_id?: string;
  procurement_batch_closed?: boolean;
  user_delivery_chat_started_at?: string;
}

interface ProcurementBatch {
  id: string;
  name: string;
  table_ids: string[];
  table_names?: string[];
  procurement_date_needed: string;
  created_at: string;
  created_by: string;
  po_file_url?: string;
  po_file_data?: string;
  po_file_mime?: string;
  po_file_name?: string;
  po_file_size?: number;
  po_file_type?: string;
  po_uploaded_at?: string;
  po_uploaded_by?: string;
  request_closed?: boolean;
  request_closed_at?: string;
  request_closed_by?: string;
  request_status?: 'PENDING' | 'DONE';
  close_request_pending?: boolean;
  close_request_production_accepted?: boolean;
  close_request_initiated_at?: string;
  close_request_initiated_by?: string;
  close_request_mode?: 'procurement';
  transfer_materials?: BatchTransferMaterial[];
}

interface ChatMessage {
  id: string;
  sender_id: string;
  sender_email: string;
  sender_role: string;
  message: string;
  message_type: 'text' | 'close_request' | 'close_response' | 'po_upload' | 'sm_upload' | 'image' | 'system';
  created_at: string;
  close_response?: 'accepted' | 'rejected';
  po_file_name?: string;
  image_data?: string;
  image_mime?: string;
  _source?: 'batch' | 'table';
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
  brand?: string;
  supplier?: string;
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

interface TransferMaterialPreview {
  key: string;
  raw_material: string;
  unit: string;
  type: string;
  brand: string;
  supplier: string;
  quantity: number;
  originalQuantity: number;
  originalUnit: string;
  excluded: boolean;
}

interface BatchTransferMaterial {
  raw_material: string;
  unit: string;
  type: string;
  brand?: string;
  supplier?: string;
  quantity: number;
  original_quantity: number;
  original_unit: string;
  excluded: boolean;
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
  styleUrls: ['./page3.component.css', '../../styles/dashboard-pages.enhancements.css']
})
export class Page3Component implements OnInit, OnDestroy {

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
  paginatedProcurementSummaries: ProcurementTableSummary[] = [];
  procurementTotalPages = 1;
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
  showProductionNotesModal = false;
  showTransferToProcurementModal = false;
  transferDateNeeded = '';
  transferSelectedTableIds: Set<string> = new Set();
  transferMaterialPreview: TransferMaterialPreview[] = [];
  productionNotesText = '';
  productionNotesReadOnly = false;

  tableStatusFilter: 'pending' | 'done' | 'all' = 'pending';

  procurementBatches: ProcurementBatch[] = [];
  selectedBatchId = '';
  selectedBatch: ProcurementBatch | null = null;
  procurementBatchSummaries: ProcurementTableSummary[] = [];
  procurementConsolidatedMaterials: ProcurementMaterialSummary[] = [];
  paginatedProcurementMaterials: ProcurementMaterialSummary[] = [];

  showChatPanel = false;
  chatMinimized = false;
  chatContextType: 'table' | 'batch' | 'unified' = 'table';
  chatTable: Table | null = null;
  chatBatch: ProcurementBatch | null = null;
  chatMessages: ChatMessage[] = [];
  chatInput = '';
  chatLoading = false;
  private chatUnsubscribe: Unsubscribe | null = null;
  private unifiedBatchUnsubscribe: Unsubscribe | null = null;
  unreadChatByTableId: Record<string, boolean> = {};
  unreadChatByBatchId: Record<string, boolean> = {};
  private unreadChatUnsubscribes: Unsubscribe[] = [];

  showPoViewerModal = false;
  poViewerTable: Table | ProcurementBatch | null = null;
  docViewerType: 'po' | 'sm' = 'po';

  viewMode: 'my_tables' | 'store_submissions' | 'for_delivery' | 'production_reviewed' | 'procurement_reviewed' = 'my_tables';
  selectedProductionView: 'submissions' | 'reviewed' = 'submissions';
  showAllPending = false;
  submitted = false;
  isLoading = false;
  isSubmitting = false;
  today = new Date().toISOString().split('T')[0];
  tomorrow: string = '';

  isUploadingPo = false;
  isUploadingSm = false;
  private readonly maxPoFileBytes = 500 * 1024;
  private readonly maxChatImageBytes = 500 * 1024;

  formData: any = {
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
  readonly skeletonRows = [1, 2, 3, 4, 5, 6];

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
    private emailNotificationService: EmailNotificationService,
    private sanitizer: DomSanitizer,
    private toast: ToastService
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
        await this.loadProcurementBatches();
      } else if (this.userRole === 'procurement') {
        await this.loadTablesDirectly();
        await this.loadProcurementBatches();
      } else {
        await this.loadTablesDirectly();
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
          sm_file_url: data['sm_file_url'],
          sm_file_data: data['sm_file_data'],
          sm_file_mime: data['sm_file_mime'],
          sm_file_name: data['sm_file_name'],
          sm_file_size: data['sm_file_size'],
          sm_file_type: data['sm_file_type'],
          sm_uploaded_at: data['sm_uploaded_at'],
          sm_uploaded_by: data['sm_uploaded_by'],
          production_reviewed: data['production_reviewed'] || false,
          production_reviewed_at: data['production_reviewed_at'],
          production_reviewed_by: data['production_reviewed_by'],
          procurement_date_needed: data['procurement_date_needed'],
          request_closed: data['request_closed'] || false,
          request_closed_at: data['request_closed_at'],
          request_closed_by: data['request_closed_by'],
          request_status: data['request_status'] || (data['request_closed'] ? 'DONE' : 'PENDING'),
          close_request_pending: data['close_request_pending'] || false,
          close_request_user_accepted: data['close_request_user_accepted'] || false,
          close_request_production_accepted: data['close_request_production_accepted'] || false,
          close_request_initiated_at: data['close_request_initiated_at'],
          close_request_initiated_by: data['close_request_initiated_by'],
          close_request_mode: data['close_request_mode'],
          procurement_batch_id: data['procurement_batch_id'],
          procurement_batch_closed: data['procurement_batch_closed'] || false,
          user_delivery_chat_started_at: data['user_delivery_chat_started_at']
        };

        if (table.user_id) {
          const emailPromise = this.getUserEmail(table.user_id).then(email => {
            table.user_email = email;
          });
          userEmailPromises.push(emailPromise);
        }

        loadedTables.push(this.sanitizeTableForRole(table));
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

      if (this.userRole === 'procurement') {
        await this.loadProcurementReviewed();
        if (this.selectedBatchId) {
          this.computeProcurementBatchSummaries();
        } else {
          this.procurementBatchSummaries = [];
          this.procurementConsolidatedMaterials = [];
          this.filteredRequisitions = [];
        }
        this.updatePagination();
      }
    } catch (err) {
      this.showToast('Failed to load tables', 'error');
    } finally {
      this.isLoading = false;
      this.setupUnreadChatListeners();
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
            production_action: existingActions[mat.raw_material.toLowerCase()]?.production_action,
            procurement_action: existingActions[mat.raw_material.toLowerCase()]?.procurement_action
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

  isTableDone(table: Table | null | undefined): boolean {
    if (!table) return false;
    return table.request_status === 'DONE' || !!table.request_closed;
  }

  isBatchDone(batch: ProcurementBatch | null | undefined): boolean {
    if (!batch) return false;
    return batch.request_status === 'DONE' || !!batch.request_closed;
  }

  getFilteredTables(): Table[] {
    if (this.tableStatusFilter === 'all') return this.tables;
    if (this.tableStatusFilter === 'done') {
      return this.tables.filter(t => this.isTableDone(t));
    }
    return this.tables.filter(t => !this.isTableDone(t));
  }

  getFilteredBatches(): ProcurementBatch[] {
    if (this.tableStatusFilter === 'all') return this.procurementBatches;
    if (this.tableStatusFilter === 'done') {
      return this.procurementBatches.filter(b => this.isBatchDone(b));
    }
    return this.procurementBatches.filter(b => !this.isBatchDone(b));
  }

  setTableStatusFilter(filter: 'pending' | 'done' | 'all') {
    this.tableStatusFilter = filter;
    if (this.selectedTable && !this.getFilteredTables().some(t => t.id === this.selectedTable!.id)) {
      this.selectedTableId = '';
      this.selectedTable = null;
      this.filteredRequisitions = [];
    }
    if (this.selectedBatch && !this.getFilteredBatches().some(b => b.id === this.selectedBatch!.id)) {
      this.selectedBatchId = '';
      this.selectedBatch = null;
      this.procurementBatchSummaries = [];
      this.procurementConsolidatedMaterials = [];
    }
  }

  async loadProcurementBatches() {
    try {
      const batchesRef = collection(this.firestore, 'procurement_batches');
      const snapshot = await this.run(() => getDocs(query(batchesRef, orderBy('created_at', 'desc'))));
      const batches: ProcurementBatch[] = [];
      snapshot.forEach(docSnap => {
        const data = docSnap.data();
        batches.push({
          id: docSnap.id,
          name: data['name'] || 'Procurement Batch',
          table_ids: data['table_ids'] || [],
          table_names: data['table_names'] || [],
          procurement_date_needed: data['procurement_date_needed'] || '',
          created_at: data['created_at'] || '',
          created_by: data['created_by'] || '',
          po_file_url: data['po_file_url'],
          po_file_data: data['po_file_data'],
          po_file_mime: data['po_file_mime'],
          po_file_name: data['po_file_name'],
          po_file_size: data['po_file_size'],
          po_file_type: data['po_file_type'],
          po_uploaded_at: data['po_uploaded_at'],
          po_uploaded_by: data['po_uploaded_by'],
          request_closed: data['request_closed'] || false,
          request_closed_at: data['request_closed_at'],
          request_closed_by: data['request_closed_by'],
          request_status: data['request_status'] || (data['request_closed'] ? 'DONE' : 'PENDING'),
          close_request_pending: data['close_request_pending'] || false,
          close_request_production_accepted: data['close_request_production_accepted'] || false,
          close_request_initiated_at: data['close_request_initiated_at'],
          close_request_initiated_by: data['close_request_initiated_by'],
          close_request_mode: data['close_request_mode'],
          transfer_materials: data['transfer_materials'] || []
        });
      });
      this.procurementBatches = batches;
    } catch (err) {
      console.error('Failed to load procurement batches', err);
    }
  }

  getBatchById(batchId: string): ProcurementBatch | null {
    return this.procurementBatches.find(b => b.id === batchId) || null;
  }

  isTableInActiveProcurementBatch(table: Table | null | undefined): boolean {
    if (!table?.production_reviewed || !table.procurement_batch_id) return false;
    return !table.procurement_batch_closed;
  }

  isTableInUserDeliveryPhase(table: Table | null | undefined): boolean {
    return !!table?.production_reviewed && !!table.procurement_batch_closed && !this.isTableDone(table);
  }

  toggleTransferSelection(table: Table, event?: Event) {
    event?.stopPropagation();
    if (!this.isTableReadyForTransfer(table)) return;
    if (this.transferSelectedTableIds.has(table.id)) {
      this.transferSelectedTableIds.delete(table.id);
    } else {
      this.transferSelectedTableIds.add(table.id);
    }
    this.transferSelectedTableIds = new Set(this.transferSelectedTableIds);
  }

  isTransferSelected(table: Table): boolean {
    return this.transferSelectedTableIds.has(table.id);
  }

  getTransferSelectedTables(): Table[] {
    return this.tables.filter(t => this.transferSelectedTableIds.has(t.id));
  }

  isTableReadyForTransfer(table: Table): boolean {
    if (this.userRole !== 'production' || table.production_reviewed) return false;
    return this.productionSubmissions.some(r => r.table_id === table.id);
  }

  canTransferSelectedTables(): boolean {
    return this.getTransferSelectedTables().length > 0 &&
      this.getTransferSelectedTables().every(t => this.isTableReadyForTransfer(t));
  }

  private getTransferMaterialKey(rawMaterial: string, unit: string, type: string): string {
    return `${rawMaterial.toLowerCase()}|${unit}|${type}`;
  }

  buildTransferMaterialPreview() {
    const tables = this.getTransferSelectedTables();
    const map = new Map<string, TransferMaterialPreview>();

    for (const table of tables) {
      const reqs = this.productionSubmissions.filter(r => r.table_id === table.id);
      for (const req of reqs) {
        if (!req.materials?.length) {
          const skuCode = String(req.skuCode ?? req['sku_code'] ?? '').trim();
          const skuName = String(req.skuName ?? req['sku_name'] ?? '').trim();
          req.materials = this.getMaterialsFromLoadedData(skuCode, skuName);
        }

        for (const mat of req.materials || []) {
          const unit = mat.unit || '—';
          const type = mat.type || 'N/A';
          const key = this.getTransferMaterialKey(mat.raw_material, unit, type);
          const qty = this.calculateMaterialTotal(
            req.quantity || req['qty_needed'] || 0,
            mat.quantity_per_batch
          );

          const existing = map.get(key);
          if (!existing) {
            map.set(key, {
              key,
              raw_material: mat.raw_material,
              unit,
              type,
              brand: '',
              supplier: '',
              quantity: qty,
              originalQuantity: qty,
              originalUnit: unit,
              excluded: false
            });
          } else {
            existing.quantity += qty;
            existing.originalQuantity += qty;
          }
        }
      }
    }

    this.transferMaterialPreview = Array.from(map.values()).sort((a, b) =>
      a.raw_material.localeCompare(b.raw_material)
    );
  }

  toggleTransferMaterialExcluded(item: TransferMaterialPreview) {
    item.excluded = !item.excluded;
  }

  getActiveTransferMaterialCount(): number {
    return this.transferMaterialPreview.filter(m => !m.excluded).length;
  }

  canConfirmTransfer(): boolean {
    return !!this.transferDateNeeded && this.getActiveTransferMaterialCount() > 0 && !this.isSubmitting;
  }

  private getTransferPreviewMap(): Map<string, TransferMaterialPreview> {
    return new Map(this.transferMaterialPreview.map(m => [m.key, m]));
  }

  getChatContextForTable(table: Table): { type: 'table' | 'batch'; batchId?: string } {
    if (
      this.isTableInActiveProcurementBatch(table) &&
      (this.userRole === 'production' || this.userRole === 'procurement' || this.userRole === 'admin')
    ) {
      return { type: 'batch', batchId: table.procurement_batch_id };
    }
    return { type: 'table' };
  }

  getChatTitle(): string {
    if (this.chatContextType === 'batch' && this.chatBatch) {
      const names = this.chatBatch.table_names?.length
        ? this.chatBatch.table_names.join(', ')
        : `${this.chatBatch.table_ids.length} tables`;
      return `${this.chatBatch.name} (${names})`;
    }
    if (this.chatContextType === 'unified' && this.chatTable) {
      return `${this.chatTable.name} — Delivery`;
    }
    return this.chatTable?.name || '';
  }

  shouldUseUnifiedTableChat(table: Table): boolean {
    return !!table.procurement_batch_id && !!table.procurement_batch_closed;
  }

  async selectBatch(batch: ProcurementBatch) {
    this.selectedBatchId = batch.id;
    this.selectedBatch = batch;
    this.selectedTableId = '';
    this.selectedTable = null;
    this.showTableDropdown = false;
    this.computeProcurementBatchSummaries();
    this.currentPage = 1;
    this.updatePagination();
  }

  computeProcurementBatchSummaries() {
    if (!this.selectedBatch) {
      this.procurementBatchSummaries = [];
      this.procurementConsolidatedMaterials = [];
      this.updateProcurementPagination();
      return;
    }

    const batchTransferMaterials = this.selectedBatch.transfer_materials || [];
    if (batchTransferMaterials.length > 0) {
      this.procurementConsolidatedMaterials = batchTransferMaterials
        .filter(tm => !tm.excluded)
        .map(tm => ({
          raw_material: tm.raw_material,
          unit: tm.unit || '—',
          type: tm.type || 'N/A',
          brand: (tm.brand || '').trim(),
          supplier: (tm.supplier || '').trim(),
          totalQuantity: tm.quantity,
          table_id: 'consolidated',
          table_name: this.selectedBatch!.name,
          production_status: 'confirmed' as const
        }));
      this.procurementBatchSummaries = [];
      this.updateProcurementPagination();
      return;
    }

    const tableIds = new Set(this.selectedBatch.table_ids);
    const reqs = this.procurementReviewed.filter(r => r.table_id && tableIds.has(r.table_id));
    const consolidatedMap = new Map<string, ProcurementMaterialSummary>();

    for (const req of reqs) {
      if (!req.table_id) continue;
      if (req.status === 'Removed' || req.production_action === 'removed') continue;

      if (!req.materials?.length) {
        const skuCode = String(req.skuCode ?? req['sku_code'] ?? '').trim();
        const skuName = String(req.skuName ?? req['sku_name'] ?? '').trim();
        req.materials = this.getMaterialsFromLoadedData(skuCode, skuName);
      }

      for (const mat of req.materials || []) {
        if (mat.production_action === 'removed') continue;

        const qty = this.calculateMaterialTotal(req.quantity || req['qty_needed'] || 0, mat.quantity_per_batch);
        const key = `${mat.raw_material.toLowerCase()}|${mat.unit}|${mat.type}`;

        let existing = consolidatedMap.get(key);
        if (!existing) {
          consolidatedMap.set(key, {
            raw_material: mat.raw_material,
            unit: mat.unit || '—',
            type: mat.type || 'N/A',
            totalQuantity: qty,
            table_id: 'consolidated',
            table_name: this.selectedBatch!.name,
            production_status: mat.production_action ?? null
          });
        } else {
          existing.totalQuantity += qty;
        }
      }
    }

    this.procurementConsolidatedMaterials = Array.from(consolidatedMap.values()).sort((a, b) =>
      a.raw_material.localeCompare(b.raw_material)
    );
    this.procurementBatchSummaries = [];
    this.updateProcurementPagination();
  }

  showRequisitionTableColumn(): boolean {
    return (this.userRole === 'production' || this.userRole === 'procurement') && !this.selectedTable;
  }

  getColspan(): number {
    let cols = 1;
    if (this.showRequisitionTableColumn()) cols++;
    if (this.userRole === 'production') cols++;
    cols += 9;
    if (this.userRole !== 'production') cols += 2;
    return cols;
  }

  getProcurementSummaryColspan(): number {
    return 5;
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
        sm_file_url: data['sm_file_url'],
        sm_file_data: data['sm_file_data'],
        sm_file_mime: data['sm_file_mime'],
        sm_file_name: data['sm_file_name'],
        sm_file_size: data['sm_file_size'],
        sm_file_type: data['sm_file_type'],
        sm_uploaded_at: data['sm_uploaded_at'],
        sm_uploaded_by: data['sm_uploaded_by'],
        production_reviewed: data['production_reviewed'] || false,
        production_reviewed_at: data['production_reviewed_at'],
        production_reviewed_by: data['production_reviewed_by'],
        procurement_date_needed: data['procurement_date_needed'],
        request_closed: data['request_closed'] || false,
        request_closed_at: data['request_closed_at'],
        request_closed_by: data['request_closed_by'],
        procurement_batch_id: data['procurement_batch_id'],
        procurement_batch_closed: data['procurement_batch_closed'] || false
      };

      if (table.user_id) {
        table.user_email = await this.getUserEmail(table.user_id);
      }

      return this.sanitizeTableForRole(table);
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
        this.selectedTable = this.sanitizeTableForRole({
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
          sm_file_url: data['sm_file_url'],
          sm_file_data: data['sm_file_data'],
          sm_file_mime: data['sm_file_mime'],
          sm_file_name: data['sm_file_name'],
          sm_file_size: data['sm_file_size'],
          sm_file_type: data['sm_file_type'],
          request_closed: data['request_closed'] || false,
          request_closed_at: data['request_closed_at'],
          request_closed_by: data['request_closed_by'],
          request_status: data['request_status'] || (data['request_closed'] ? 'DONE' : 'PENDING'),
          close_request_pending: data['close_request_pending'] || false,
          close_request_user_accepted: data['close_request_user_accepted'] || false,
          close_request_production_accepted: data['close_request_production_accepted'] || false,
          close_request_initiated_at: data['close_request_initiated_at'],
          close_request_initiated_by: data['close_request_initiated_by'],
          close_request_mode: data['close_request_mode']
        });
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
  // Request Delivered (via Chat)
  async closeRequest(table: Table) {
    this.openChat(table);
    if (this.userRole === 'procurement' || this.userRole === 'production' || this.userRole === 'admin') {
      await this.initiateCloseRequestInChat();
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

  async submitReviewedTables() {
    let tablesToTransfer = this.getTransferSelectedTables();
    if (tablesToTransfer.length === 0 && this.selectedTable && this.isTableReadyForTransfer(this.selectedTable)) {
      tablesToTransfer = [this.selectedTable];
    }
    if (tablesToTransfer.length === 0 || !this.transferDateNeeded) return;
    if (this.getActiveTransferMaterialCount() === 0) {
      this.showToast('Include at least one raw material in the transfer', 'error');
      return;
    }

    try {
      this.isSubmitting = true;
      this.isLoading = true;

      const batchName = `Procurement Batch ${this.formatTableDate(new Date())}`;
      const tableIds = tablesToTransfer.map(t => t.id);
      const tableNames = tablesToTransfer.map(t => t.name);
      const now = new Date().toISOString();
      const excludedMaterialNames = new Set(
        this.transferMaterialPreview.filter(p => p.excluded).map(p => p.raw_material.toLowerCase())
      );
      const transferMaterials: BatchTransferMaterial[] = this.transferMaterialPreview.map(m => ({
        raw_material: m.raw_material,
        unit: m.unit,
        type: m.type,
        brand: (m.brand || '').trim(),
        supplier: (m.supplier || '').trim(),
        quantity: m.quantity,
        original_quantity: m.originalQuantity,
        original_unit: m.originalUnit,
        excluded: m.excluded
      }));

      const batchDocRef = await this.run(() =>
        addDoc(collection(this.firestore, 'procurement_batches'), {
          name: batchName,
          table_ids: tableIds,
          table_names: tableNames,
          procurement_date_needed: this.transferDateNeeded,
          transfer_materials: transferMaterials,
          created_at: now,
          created_by: this.userId,
          request_status: 'PENDING',
          request_closed: false
        })
      );

      let totalItems = 0;
      let confirmedItems = 0;
      let removedItems = 0;

      for (const table of tablesToTransfer) {
        const tableSubmissions = this.productionSubmissions.filter(r => r.table_id === table.id);
        totalItems += tableSubmissions.length;

        const updatePromises = tableSubmissions.map(async (req) => {
          if (!req.materials?.length) {
            const skuCode = String(req.skuCode ?? req['sku_code'] ?? '').trim();
            const skuName = String(req.skuName ?? req['sku_name'] ?? '').trim();
            req.materials = this.getMaterialsFromLoadedData(skuCode, skuName);
          }

          const materials = (req.materials || []).map(mat => {
            const excluded = excludedMaterialNames.has(mat.raw_material.toLowerCase());
            return {
              ...mat,
              production_action: excluded ? 'removed' as const : 'confirmed' as const
            };
          });

          const allRemoved = materials.length > 0 && materials.every(m => m.production_action === 'removed');
          const action = allRemoved ? 'removed' : 'confirmed';
          const newStatus = action === 'removed' ? 'Removed' : 'Production_Confirmed';

          if (action === 'confirmed') confirmedItems++;
          else removedItems++;

          const updateData: any = {
            materials,
            production_action: action,
            production_action_at: now,
            production_action_by: this.userId
          };

          return this.db.updateRequisitionStatus(req.id, newStatus, this.userId, table.id, updateData);
        });

        await Promise.all(updatePromises);

        const tableRef = doc(this.firestore, 'tables', table.id);
        await this.run(() =>
          updateDoc(tableRef, {
            production_reviewed: true,
            production_reviewed_at: now,
            production_reviewed_by: this.userId,
            procurement_date_needed: this.transferDateNeeded,
            procurement_batch_id: batchDocRef.id,
            procurement_batch_closed: false,
            updated_at: now,
            submitted: true
          })
        );

        table.production_reviewed = true;
        table.procurement_date_needed = this.transferDateNeeded;
        table.procurement_batch_id = batchDocRef.id;
        table.procurement_batch_closed = false;
        const tableIndex = this.tables.findIndex(t => t.id === table.id);
        if (tableIndex !== -1) {
          this.tables[tableIndex] = { ...this.tables[tableIndex], ...table };
        }

        await this.notificationService.sendTableReviewedByProductionNotification(
          table.id,
          table.name,
          this.userId
        );
      }

      const currentUser = await this.auth.getCurrentUserPromise();
      const reviewerEmail = currentUser?.email || 'unknown@example.com';

      try {
        await this.emailNotificationService.sendTableReviewedNotification({
          tableName: batchName,
          reviewerEmail,
          reviewedAt: now,
          totalItems,
          confirmedItems,
          removedItems,
          tableId: batchDocRef.id,
          reviewLink: `${window.location.origin}/requisitions?batchId=${batchDocRef.id}&role=procurement`
        });
        this.showToast(
          `${tablesToTransfer.length} table(s) transferred to procurement as one consolidated batch`,
          'success'
        );
      } catch {
        this.showToast(
          `${tablesToTransfer.length} table(s) transferred (email notification failed)`,
          'info'
        );
      }

      this.transferSelectedTableIds = new Set();
      this.closeTransferToProcurementModal();
      await this.loadProductionSubmissions();
      await this.loadProductionReviewed();
      await this.loadProcurementBatches();

      if (this.selectedTable) {
        this.selectedProductionView = 'reviewed';
        this.filteredRequisitions = this.productionReviewed.filter(r => r.table_id === this.selectedTable!.id);
      }
      this.updatePagination();
      this.setupUnreadChatListeners();
    } catch (err) {
      this.showToast('Failed to submit tables to procurement', 'error');
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
    if (this.isProductionActionsLocked(req)) {
      this.showToast('Production actions are locked after transfer to Procurement', 'error');
      return;
    }

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
    if (this.isProductionActionsLocked(req)) {
      this.showToast('Production actions are locked after transfer to Procurement', 'error');
      return;
    }

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
    this.updateProcurementPagination();
  }

  setProcurementMaterialAction(item: ProcurementMaterialSummary, action: 'approved' | 'rejected') {
    const table = this.tables.find(t => t.id === item.table_id);
    if (this.isRequestDone(table)) {
      this.showToast('Request is DONE — procurement actions are locked', 'error');
      return;
    }
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

  openProcurementOriginalItemsModalForSelectedTable() {
    if (!this.selectedTable) return;
    const summary = this.procurementTableSummaries.find(s => s.table_id === this.selectedTable!.id);
    if (summary) {
      this.openProcurementOriginalItemsModal(summary);
    }
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

  hasPoDocument(table: Table | ProcurementBatch | null | undefined): boolean {
    if (!table) return false;
    return !!(table.po_file_data || table.po_file_url);
  }

  hasSmDocument(table: Table | null | undefined): boolean {
    if (!table) return false;
    return !!(table.sm_file_data || table.sm_file_url);
  }

  isProcurementInWorkflow(table: Table | null | undefined): boolean {
    if (this.chatContextType === 'batch') return true;
    return !!table?.production_reviewed && !this.isTableInActiveProcurementBatch(table);
  }

  isBatchProcurementWorkflow(): boolean {
    return this.chatContextType === 'batch' && !!this.chatBatch && !this.isBatchDone(this.chatBatch);
  }

  getCloseRequestMode(table: Table | null | undefined): 'procurement' | 'production' {
    if (table?.close_request_mode) return table.close_request_mode;
    return this.isProcurementInWorkflow(table) ? 'procurement' : 'production';
  }

  getDeliveryRequestBannerText(): string {
    if (this.chatContextType === 'batch') {
      return 'Procurement requested batch delivery confirmation. Production must accept or reject:';
    }
    const mode = this.getCloseRequestMode(this.chatTable);
    const initiator = mode === 'procurement' ? 'Procurement' : 'Production';
    return `${initiator} requested delivery confirmation. Please accept or reject:`;
  }

  getChatStatusLabel(): 'PENDING' | 'DONE' {
    if (this.chatContextType === 'batch' && this.chatBatch) {
      return this.isBatchDone(this.chatBatch) ? 'DONE' : 'PENDING';
    }
    return this.getRequestStatusLabel(this.chatTable);
  }

  canShowSmInChat(): boolean {
    return this.canViewSm() && !!this.chatTable && this.hasSmDocument(this.chatTable);
  }

  private sanitizeTableForRole(table: Table): Table {
    if (this.userRole === 'user' || this.userRole === 'store') {
      return {
        ...table,
        po_file_url: undefined,
        po_file_data: undefined,
        po_file_mime: undefined,
        po_file_name: undefined,
        po_file_size: undefined,
        po_file_type: undefined
      };
    }
    return table;
  }

  canViewPo(): boolean {
    if (this.chatContextType === 'batch' || this.chatContextType === 'unified') {
      return this.userRole === 'procurement' || this.userRole === 'production' || this.userRole === 'admin';
    }
    if (!this.isProcurementInWorkflow(this.chatTable)) return false;
    return this.userRole === 'procurement' || this.userRole === 'production' || this.userRole === 'admin';
  }

  canViewSm(): boolean {
    if (this.chatContextType === 'batch') return false;
    return this.userRole === 'user' || this.userRole === 'store' ||
      this.userRole === 'production' || this.userRole === 'procurement' || this.userRole === 'admin';
  }

  shouldShowChatMessage(msg: ChatMessage): boolean {
    const source: 'batch' | 'table' = msg._source || (this.chatContextType === 'batch' ? 'batch' : 'table');

    if (this.userRole === 'production' || this.userRole === 'admin') {
      if (msg.message_type === 'po_upload' && source === 'table' && !this.isProcurementInWorkflow(this.chatTable)) {
        return false;
      }
      if (msg.message_type === 'po_upload' && !this.canViewPo()) return false;
      return true;
    }

    if (this.userRole === 'procurement') {
      return source === 'batch';
    }

    if (this.userRole === 'user' || this.userRole === 'store') {
      if (source === 'batch') return false;
      const startedAt = this.chatTable?.user_delivery_chat_started_at;
      if (startedAt && msg.created_at < startedAt) return false;
      if (msg.message_type === 'po_upload') return false;
      if (msg.sender_role === 'procurement') return false;
      return true;
    }

    if (msg.message_type === 'po_upload' && !this.canViewPo()) return false;
    return true;
  }

  shouldShowChatPhaseDivider(msg: ChatMessage, index: number): boolean {
    if (this.chatContextType !== 'unified') return false;
    if (this.userRole !== 'production' && this.userRole !== 'admin') return false;
    const source: 'batch' | 'table' = msg._source || 'table';
    if (source !== 'table' || !this.shouldShowChatMessage(msg)) return false;

    for (let i = index - 1; i >= 0; i--) {
      const prev = this.chatMessages[i];
      if (!this.shouldShowChatMessage(prev)) continue;
      const prevSource: 'batch' | 'table' = prev._source || 'table';
      return prevSource === 'batch';
    }
    return false;
  }

  getPoViewerTarget(): Table | ProcurementBatch | null {
    if ((this.chatContextType === 'batch' || this.chatContextType === 'unified') && this.chatBatch) {
      return this.chatBatch;
    }
    return this.chatTable;
  }

  canShowPoInChat(): boolean {
    if (this.chatContextType === 'batch' || this.chatContextType === 'unified') {
      return this.canViewPo() && !!this.chatBatch && this.hasPoDocument(this.chatBatch);
    }
    return this.canViewPo() && !!this.chatTable && this.hasPoDocument(this.chatTable);
  }

  getDocViewUrl(table: Table | ProcurementBatch | null | undefined, type: 'po' | 'sm'): string {
    if (!table) return '';
    if (type === 'sm') {
      const t = table as Table;
      const data = t.sm_file_data;
      const mime = t.sm_file_mime || t.sm_file_type;
      const url = t.sm_file_url;
      if (data) {
        return `data:${mime || 'application/octet-stream'};base64,${data}`;
      }
      return url || '';
    }
    const data = table.po_file_data;
    const mime = table.po_file_mime || table.po_file_type;
    const url = table.po_file_url;
    if (data) {
      return `data:${mime || 'application/octet-stream'};base64,${data}`;
    }
    return url || '';
  }

  getSafeDocViewUrl(table: Table | ProcurementBatch | null | undefined, type: 'po' | 'sm'): SafeResourceUrl | string {
    const url = this.getDocViewUrl(table, type);
    return url ? this.sanitizer.bypassSecurityTrustResourceUrl(url) : '';
  }

  getDocMime(table: Table | ProcurementBatch | null | undefined, type: 'po' | 'sm'): string {
    if (!table) return '';
    if (type === 'po') {
      return table.po_file_mime || table.po_file_type || 'application/octet-stream';
    }
    const t = table as Table;
    return t.sm_file_mime || t.sm_file_type || 'application/octet-stream';
  }

  isDocImage(table: Table | ProcurementBatch | null | undefined, type: 'po' | 'sm'): boolean {
    return this.getDocMime(table, type).startsWith('image/');
  }

  isDocPdf(table: Table | ProcurementBatch | null | undefined, type: 'po' | 'sm'): boolean {
    const mime = this.getDocMime(table, type).toLowerCase();
    const name = ((type === 'po' ? table?.po_file_name : (table as Table)?.sm_file_name) || '').toLowerCase();
    return mime.includes('pdf') || name.endsWith('.pdf');
  }

  getDocFileName(table: Table | ProcurementBatch | null | undefined, type: 'po' | 'sm'): string {
    if (!table) return type === 'po' ? 'Purchase Order' : 'Shipping Manifest';
    if (type === 'po') {
      return table.po_file_name || 'Purchase Order';
    }
    return (table as Table).sm_file_name || 'Shipping Manifest';
  }

  getPoViewUrl(table: Table | null | undefined): string {
    return this.getDocViewUrl(table, 'po');
  }

  getSafePoViewUrl(table: Table | null | undefined): SafeResourceUrl | string {
    return this.getSafeDocViewUrl(table, 'po');
  }

  getPoMime(table: Table | null | undefined): string {
    return this.getDocMime(table, 'po');
  }

  isPoImage(table: Table | null | undefined): boolean {
    return this.isDocImage(table, 'po');
  }

  isPoPdf(table: Table | null | undefined): boolean {
    return this.isDocPdf(table, 'po');
  }

  getChatImageUrl(msg: ChatMessage): string {
    if (!msg.image_data) return '';
    return `data:${msg.image_mime || 'image/jpeg'};base64,${msg.image_data}`;
  }

  getSafeChatImageUrl(msg: ChatMessage): SafeResourceUrl | string {
    const url = this.getChatImageUrl(msg);
    return url ? this.sanitizer.bypassSecurityTrustResourceUrl(url) : '';
  }

  openPoViewer(table: Table | null | undefined) {
    this.openDocViewer(table, 'po');
  }

  openDocViewer(table: Table | ProcurementBatch | null | undefined, type: 'po' | 'sm') {
    if (!table) return;
    if (type === 'po' && (!this.canViewPo() || !this.hasPoDocument(table))) return;
    if (type === 'sm' && (!(table as Table).sm_file_data && !(table as Table).sm_file_url)) return;
    if (type === 'sm' && (!this.canViewSm() || !this.hasSmDocument(table as Table))) return;
    this.poViewerTable = table;
    this.docViewerType = type;
    this.showPoViewerModal = true;
  }

  closePoViewerModal() {
    this.showPoViewerModal = false;
    this.poViewerTable = null;
  }

  private getTableMessagesRef(tableId: string) {
    return collection(this.firestore, 'tables', tableId, 'messages');
  }

  private getBatchMessagesRef(batchId: string) {
    return collection(this.firestore, 'procurement_batches', batchId, 'messages');
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
    return this.canTransferSelectedTables();
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
    return this.userRole === 'production' && req.status === 'Submitted' && !this.isProductionActionsLocked(req);
  }

  canProductionEditMaterials(req: Requisition): boolean {
    return false;
  }

  getTableForRequisition(req: Requisition): Table | null {
    if (req.table_id) {
      const found = this.tables.find(t => t.id === req.table_id);
      if (found) return found;
    }
    if (this.selectedTable && this.selectedTable.id === req.table_id) {
      return this.selectedTable;
    }
    return null;
  }

  isProductionActionsLocked(req: Requisition): boolean {
    const table = this.getTableForRequisition(req);
    return !!table?.production_reviewed;
  }

  isRequestDone(table: Table | null | undefined): boolean {
    if (!table) return false;
    return table.request_status === 'DONE' || !!table.request_closed;
  }

  getRequestStatusLabel(table: Table | null | undefined): 'PENDING' | 'DONE' {
    return this.isRequestDone(table) ? 'DONE' : 'PENDING';
  }

  canOpenChat(table: Table | null | undefined): boolean {
    if (!table) return false;
    if (this.userRole === 'admin') return true;
    if ((this.userRole === 'user' || this.userRole === 'store') && table.user_id === this.userId) {
      return !!table.submitted;
    }
    if (this.userRole === 'production') return !!table.submitted;
    if (this.userRole === 'procurement') return !!table.production_reviewed;
    return false;
  }

  canOpenBatchChat(batch: ProcurementBatch | null | undefined): boolean {
    if (!batch) return false;
    return this.userRole === 'production' || this.userRole === 'procurement' || this.userRole === 'admin';
  }

  openChatForCurrentContext() {
    if (this.userRole === 'procurement' && this.selectedBatch) {
      this.openBatchChat(this.selectedBatch);
      return;
    }
    if (this.selectedTable) {
      this.openChat(this.selectedTable);
    }
  }

  hasChatDocuments(): boolean {
    return this.canShowPoInChat() || this.canShowSmInChat();
  }

  hasChatWritableToolbar(): boolean {
    return this.isChatWritable() && (
      this.canInitiateCloseRequest() || this.canUploadPo() || this.canUploadSm()
    );
  }

  showChatToolbar(): boolean {
    return this.hasChatDocuments() || this.hasChatWritableToolbar();
  }

  getPageSubtitle(): string {
    if (this.userRole === 'production') {
      return this.selectedTable
        ? `Reviewing ${this.selectedTable.name}`
        : 'Review submitted requisition tables';
    }
    if (this.userRole === 'procurement') {
      return 'Production-confirmed tables ready for delivery';
    }
    if (this.userRole === 'admin') {
      return 'Full workflow access';
    }
    return 'Create tables and track material requests';
  }

  getTableStatLabel(): string {
    const count = this.filteredRequisitions.length;
    if (this.userRole === 'production') return `${count} to review`;
    return `${count} requisition${count === 1 ? '' : 's'}`;
  }

  getProcurementMaterialCount(): number {
    if (this.userRole === 'procurement' && this.selectedBatch) {
      return this.procurementConsolidatedMaterials.length;
    }
    if (!this.selectedTable) return 0;
    const summary = this.procurementTableSummaries.find(s => s.table_id === this.selectedTable!.id);
    return summary?.uniqueMaterialsCount ?? 0;
  }

  showTableContextBar(): boolean {
    return !!this.selectedTable || !!this.selectedBatch;
  }

  openTransferToProcurementModal() {
    if (!this.canTransferSelectedTables()) {
      if (this.selectedTable && this.isTableReadyForTransfer(this.selectedTable)) {
        this.transferSelectedTableIds = new Set([this.selectedTable.id]);
      } else {
        this.showToast('Select one or more submitted tables to transfer', 'error');
        return;
      }
    }
    this.transferDateNeeded = '';
    this.buildTransferMaterialPreview();
    if (this.transferMaterialPreview.length === 0) {
      this.showToast('No raw materials found for the selected tables', 'error');
      return;
    }
    this.showTransferToProcurementModal = true;
  }

  closeTransferToProcurementModal() {
    this.showTransferToProcurementModal = false;
    this.transferDateNeeded = '';
    this.transferMaterialPreview = [];
  }

  confirmTransferToProcurement() {
    if (!this.transferDateNeeded) {
      this.showToast('Please set a date needed before transferring', 'error');
      return;
    }
    if (this.getActiveTransferMaterialCount() === 0) {
      this.showToast('Include at least one raw material in the transfer', 'error');
      return;
    }
    this.submitReviewedTables();
  }

  isChatWritable(): boolean {
    if (this.chatContextType === 'batch') {
      return !!this.chatBatch && !this.isBatchDone(this.chatBatch);
    }
    if (this.chatContextType === 'unified' || this.chatContextType === 'table') {
      if (!this.chatTable) return false;
      if (this.isTableInActiveProcurementBatch(this.chatTable)) {
        return false;
      }
      return !this.isRequestDone(this.chatTable);
    }
    return false;
  }

  canInitiateCloseRequest(): boolean {
    if (this.chatContextType === 'batch') {
      if (!this.chatBatch || this.isBatchDone(this.chatBatch) || this.chatBatch.close_request_pending) {
        return false;
      }
      if (this.userRole === 'procurement' || this.userRole === 'admin') {
        return this.hasPoDocument(this.chatBatch);
      }
      return false;
    }

    if (!this.chatTable || this.isRequestDone(this.chatTable) || this.chatTable.close_request_pending) {
      return false;
    }
    if (this.isTableInActiveProcurementBatch(this.chatTable)) {
      return false;
    }
    if (this.userRole === 'admin') return true;
    if (this.userRole === 'procurement') {
      return this.isProcurementInWorkflow(this.chatTable) && this.hasPoDocument(this.chatTable);
    }
    if (this.userRole === 'production') {
      return this.canProductionInitiateDelivery(this.chatTable);
    }
    return false;
  }

  private canProductionInitiateDelivery(table: Table): boolean {
    if (!table.submitted) return false;
    if (this.isTableInUserDeliveryPhase(table)) {
      return true;
    }
    if (table.production_reviewed) return false;
    const tableSubmissions = this.productionSubmissions.filter(r => r.table_id === table.id);
    if (tableSubmissions.length === 0) return false;
    return tableSubmissions.every(
      r => r.production_action === 'confirmed' || r.production_action === 'removed'
    );
  }

  canUploadPo(): boolean {
    if (this.chatContextType === 'batch') {
      return this.userRole === 'procurement' || this.userRole === 'admin';
    }
    return (
      (this.userRole === 'procurement' || this.userRole === 'admin') &&
      this.isProcurementInWorkflow(this.chatTable)
    );
  }

  canUploadSm(): boolean {
    if (this.chatContextType === 'batch') return false;
    if (!this.chatTable) return false;
    if (this.isTableInActiveProcurementBatch(this.chatTable)) return false;
    return (
      (this.userRole === 'production' || this.userRole === 'admin') &&
      (!this.chatTable.production_reviewed || this.isTableInUserDeliveryPhase(this.chatTable))
    );
  }

  canRespondToCloseRequest(): boolean {
    if (this.chatContextType === 'batch') {
      if (!this.chatBatch?.close_request_pending || this.isBatchDone(this.chatBatch)) return false;
      if (this.userRole === 'production') {
        return !this.chatBatch.close_request_production_accepted;
      }
      return false;
    }

    if (!this.chatTable?.close_request_pending || this.isRequestDone(this.chatTable)) return false;
    if (this.isTableInActiveProcurementBatch(this.chatTable)) return false;
    const mode = this.getCloseRequestMode(this.chatTable);
    if (this.userRole === 'user' || this.userRole === 'store') {
      return this.chatTable.user_id === this.userId && !this.chatTable.close_request_user_accepted;
    }
    if (this.userRole === 'production' && mode === 'procurement') {
      return !this.chatTable.close_request_production_accepted;
    }
    if (this.userRole === 'production' && mode === 'production') {
      return false;
    }
    return false;
  }

  openProductionNotesModal(req: Requisition, readOnly = false) {
    this.selectedRequisition = req;
    this.productionNotesText = req.production_notes || '';
    this.productionNotesReadOnly = readOnly;
    this.showProductionNotesModal = true;
  }

  closeProductionNotesModal() {
    this.showProductionNotesModal = false;
    this.selectedRequisition = null;
    this.productionNotesText = '';
    this.productionNotesReadOnly = false;
  }

  async saveProductionNotes() {
    if (!this.selectedRequisition || this.productionNotesReadOnly) return;
    if (!this.productionNotesText.trim()) {
      this.showToast('Please enter a note', 'error');
      return;
    }
    try {
      const success = await this.db.updateRequisitionStatus(
        this.selectedRequisition.id,
        this.selectedRequisition.status,
        this.userId,
        this.selectedRequisition.table_id || '',
        { production_notes: this.productionNotesText.trim() }
      );
      if (success) {
        this.selectedRequisition.production_notes = this.productionNotesText.trim();
        this.closeProductionNotesModal();
        this.showToast('Production note saved', 'success');
      } else {
        this.showToast('Failed to save note', 'error');
      }
    } catch {
      this.showToast('Failed to save note', 'error');
    }
  }

  openChat(table: Table) {
    if (!this.canOpenChat(table)) {
      this.showToast('Chat is not available for this table', 'error');
      return;
    }
    const latest = this.tables.find(t => t.id === table.id) || table;
    const ctx = this.getChatContextForTable(latest);

    if (ctx.type === 'batch' && ctx.batchId) {
      const batch = this.getBatchById(ctx.batchId);
      if (batch) {
        this.openBatchChat(batch, latest);
        return;
      }
    }

    if (this.shouldUseUnifiedTableChat(latest)) {
      this.openUnifiedTableChat(latest);
      return;
    }

    this.chatContextType = 'table';
    this.chatBatch = null;
    this.chatTable = this.sanitizeTableForRole({ ...latest });
    this.showChatPanel = true;
    this.chatMinimized = false;
    this.markChatRead(latest.id);
    this.subscribeToChatContext('table', latest.id);
  }

  openUnifiedTableChat(table: Table) {
    const latest = this.tables.find(t => t.id === table.id) || table;
    this.chatContextType = 'unified';
    this.chatTable = this.sanitizeTableForRole({ ...latest });
    this.chatBatch = null;
    if (latest.procurement_batch_id) {
      const batch = this.getBatchById(latest.procurement_batch_id);
      if (batch) {
        this.chatBatch = { ...batch };
      }
    }
    this.showChatPanel = true;
    this.chatMinimized = false;
    this.markChatRead(latest.id);
    this.subscribeToUnifiedChat(latest);
  }

  private autoOpenUserDeliveryChat(batchId: string) {
    if (this.userRole !== 'production') return;
    const batch = this.getBatchById(batchId);
    if (!batch?.table_ids.length) return;

    const preferredId =
      this.chatTable?.id && batch.table_ids.includes(this.chatTable.id)
        ? this.chatTable.id
        : this.selectedTableId && batch.table_ids.includes(this.selectedTableId)
          ? this.selectedTableId
          : batch.table_ids[0];

    const table = this.tables.find(t => t.id === preferredId);
    if (table) {
      this.openUnifiedTableChat(table);
    }
  }

  openBatchChat(batch: ProcurementBatch, contextTable?: Table) {
    if (!this.canOpenBatchChat(batch)) {
      this.showToast('Chat is not available for this batch', 'error');
      return;
    }
    const latest = this.getBatchById(batch.id) || batch;
    this.chatContextType = 'batch';
    this.chatBatch = { ...latest };
    this.chatTable = contextTable ? this.sanitizeTableForRole({ ...contextTable }) : null;
    this.showChatPanel = true;
    this.chatMinimized = false;
    this.markBatchChatRead(latest.id);
    this.subscribeToChatContext('batch', latest.id);
  }

  closeChatPanel() {
    this.showChatPanel = false;
    this.chatMinimized = false;
    if (this.chatUnsubscribe) {
      this.chatUnsubscribe();
      this.chatUnsubscribe = null;
    }
    if (this.unifiedBatchUnsubscribe) {
      this.unifiedBatchUnsubscribe();
      this.unifiedBatchUnsubscribe = null;
    }
    this.chatContextType = 'table';
    this.chatTable = null;
    this.chatBatch = null;
    this.chatMessages = [];
    this.chatInput = '';
  }

  toggleChatMinimize() {
    this.chatMinimized = !this.chatMinimized;
    if (!this.chatMinimized) {
      if (this.chatContextType === 'batch' && this.chatBatch) {
        this.markBatchChatRead(this.chatBatch.id);
      } else if (this.chatTable) {
        this.markChatRead(this.chatTable.id);
      }
    }
  }

  hasUnreadChat(table: Table | null | undefined): boolean {
    if (!table) return false;
    if (this.isTableInActiveProcurementBatch(table) && table.procurement_batch_id) {
      return !!this.unreadChatByBatchId[table.procurement_batch_id];
    }
    return !!this.unreadChatByTableId[table.id];
  }

  hasUnreadBatchChat(batch: ProcurementBatch | null | undefined): boolean {
    if (!batch) return false;
    return !!this.unreadChatByBatchId[batch.id];
  }

  private chatReadAtKey(tableId: string): string {
    return `chatReadAt_${this.userId}_${tableId}`;
  }

  private batchChatReadAtKey(batchId: string): string {
    return `chatReadAt_batch_${this.userId}_${batchId}`;
  }

  private getChatLastReadAt(tableId: string): string | null {
    return localStorage.getItem(this.chatReadAtKey(tableId));
  }

  private getBatchChatLastReadAt(batchId: string): string | null {
    return localStorage.getItem(this.batchChatReadAtKey(batchId));
  }

  private markChatRead(tableId: string) {
    localStorage.setItem(this.chatReadAtKey(tableId), new Date().toISOString());
    this.unreadChatByTableId[tableId] = false;
  }

  private markBatchChatRead(batchId: string) {
    localStorage.setItem(this.batchChatReadAtKey(batchId), new Date().toISOString());
    this.unreadChatByBatchId[batchId] = false;
  }

  private teardownUnreadChatListeners() {
    this.unreadChatUnsubscribes.forEach(unsub => unsub());
    this.unreadChatUnsubscribes = [];
  }

  private setupUnreadChatListeners() {
    this.teardownUnreadChatListeners();
    if (!this.userId) return;

    const tables = this.tables.filter(t => this.canOpenChat(t) && !this.isTableInActiveProcurementBatch(t));
    for (const table of tables) {
      const messagesRef = this.getTableMessagesRef(table.id);
      const q = query(messagesRef, orderBy('created_at', 'desc'), limit(1));
      const unsub = onSnapshot(q, (snap) => {
        if (snap.empty) {
          this.unreadChatByTableId[table.id] = false;
          return;
        }
        const msg = snap.docs[0].data();
        if (msg['sender_id'] === this.userId) {
          this.unreadChatByTableId[table.id] = false;
          return;
        }
        const isOpen = this.showChatPanel &&
          (this.chatContextType === 'table' || this.chatContextType === 'unified') &&
          this.chatTable?.id === table.id && !this.chatMinimized;
        if (isOpen) {
          this.markChatRead(table.id);
          return;
        }
        const lastRead = this.getChatLastReadAt(table.id);
        this.unreadChatByTableId[table.id] = !lastRead || msg['created_at'] > lastRead;
        this.cdr.detectChanges();
      });
      this.unreadChatUnsubscribes.push(unsub);
    }

    const batchIds = new Set<string>();
    this.tables.forEach(t => {
      if (t.procurement_batch_id && this.isTableInActiveProcurementBatch(t)) {
        batchIds.add(t.procurement_batch_id);
      }
    });
    this.procurementBatches.forEach(b => {
      if (!this.isBatchDone(b)) batchIds.add(b.id);
    });

    for (const batchId of batchIds) {
      const messagesRef = this.getBatchMessagesRef(batchId);
      const q = query(messagesRef, orderBy('created_at', 'desc'), limit(1));
      const unsub = onSnapshot(q, (snap) => {
        if (snap.empty) {
          this.unreadChatByBatchId[batchId] = false;
          return;
        }
        const msg = snap.docs[0].data();
        if (msg['sender_id'] === this.userId) {
          this.unreadChatByBatchId[batchId] = false;
          return;
        }
        const isOpen = this.showChatPanel && this.chatContextType === 'batch' &&
          this.chatBatch?.id === batchId && !this.chatMinimized;
        if (isOpen) {
          this.markBatchChatRead(batchId);
          return;
        }
        const lastRead = this.getBatchChatLastReadAt(batchId);
        this.unreadChatByBatchId[batchId] = !lastRead || msg['created_at'] > lastRead;
        this.cdr.detectChanges();
      });
      this.unreadChatUnsubscribes.push(unsub);
    }
  }

  private subscribeToChatContext(contextType: 'table' | 'batch', contextId: string) {
    if (this.unifiedBatchUnsubscribe) {
      this.unifiedBatchUnsubscribe();
      this.unifiedBatchUnsubscribe = null;
    }
    if (this.chatUnsubscribe) {
      this.chatUnsubscribe();
      this.chatUnsubscribe = null;
    }
    this.chatLoading = true;
    const messagesRef = contextType === 'batch'
      ? this.getBatchMessagesRef(contextId)
      : this.getTableMessagesRef(contextId);
    const q = query(messagesRef, orderBy('created_at', 'asc'));
    this.chatUnsubscribe = onSnapshot(q, (snap) => {
      this.chatMessages = snap.docs.map(docSnap => ({
        id: docSnap.id,
        ...(docSnap.data() as Omit<ChatMessage, 'id'>),
        _source: contextType
      }));
      this.chatLoading = false;
      if (this.showChatPanel && !this.chatMinimized) {
        if (contextType === 'batch') {
          this.markBatchChatRead(contextId);
        } else {
          this.markChatRead(contextId);
        }
      }
      this.cdr.detectChanges();
    }, () => {
      this.chatLoading = false;
    });
  }

  private subscribeToUnifiedChat(table: Table) {
    if (this.chatUnsubscribe) {
      this.chatUnsubscribe();
      this.chatUnsubscribe = null;
    }
    if (this.unifiedBatchUnsubscribe) {
      this.unifiedBatchUnsubscribe();
      this.unifiedBatchUnsubscribe = null;
    }

    this.chatLoading = true;
    let batchMessages: ChatMessage[] = [];
    let tableMessages: ChatMessage[] = [];
    let batchLoaded = !table.procurement_batch_id;
    let tableLoaded = false;

    const mergeMessages = () => {
      if (!batchLoaded || !tableLoaded) return;
      this.chatMessages = [
        ...batchMessages.map(m => ({ ...m, _source: 'batch' as const })),
        ...tableMessages.map(m => ({ ...m, _source: 'table' as const }))
      ].sort((a, b) => a.created_at.localeCompare(b.created_at));
      this.chatLoading = false;
      if (this.showChatPanel && !this.chatMinimized) {
        this.markChatRead(table.id);
      }
      this.cdr.detectChanges();
    };

    const tableQ = query(this.getTableMessagesRef(table.id), orderBy('created_at', 'asc'));
    this.chatUnsubscribe = onSnapshot(tableQ, (snap) => {
      tableMessages = snap.docs.map(docSnap => ({
        id: docSnap.id,
        ...(docSnap.data() as Omit<ChatMessage, 'id'>)
      }));
      tableLoaded = true;
      mergeMessages();
    }, () => {
      tableLoaded = true;
      mergeMessages();
    });

    if (table.procurement_batch_id) {
      const batchQ = query(this.getBatchMessagesRef(table.procurement_batch_id), orderBy('created_at', 'asc'));
      this.unifiedBatchUnsubscribe = onSnapshot(batchQ, (snap) => {
        batchMessages = snap.docs.map(docSnap => ({
          id: docSnap.id,
          ...(docSnap.data() as Omit<ChatMessage, 'id'>)
        }));
        batchLoaded = true;
        mergeMessages();
      }, () => {
        batchLoaded = true;
        mergeMessages();
      });
    }
  }

  async sendChatMessage() {
    if (!this.chatInput.trim() || !this.isChatWritable()) return;
    try {
      const messagesRef = this.chatContextType === 'batch' && this.chatBatch
        ? this.getBatchMessagesRef(this.chatBatch.id)
        : this.chatTable
          ? this.getTableMessagesRef(this.chatTable.id)
          : null;
      if (!messagesRef) return;

      await this.run(() => addDoc(messagesRef, {
        sender_id: this.userId,
        sender_email: this.userName,
        sender_role: this.userRole,
        message: this.chatInput.trim(),
        message_type: 'text',
        created_at: new Date().toISOString()
      }));
      this.chatInput = '';
    } catch (err) {
      console.error('sendChatMessage failed', err);
      this.showToast('Failed to send message. Deploy latest Firestore rules if this persists.', 'error');
    }
  }

  async initiateCloseRequestInChat() {
    if (!this.canInitiateCloseRequest()) return;

    if (this.chatContextType === 'batch' && this.chatBatch) {
      if (!confirm('Mark batch as delivered? Production must accept to confirm procurement delivery.')) return;

      const batchRef = doc(this.firestore, 'procurement_batches', this.chatBatch.id);
      const initiatedAt = new Date().toISOString();
      try {
        await this.run(() => updateDoc(batchRef, {
          close_request_pending: true,
          close_request_initiated_at: initiatedAt,
          close_request_initiated_by: this.userId,
          close_request_mode: 'procurement',
          close_request_production_accepted: false,
          updated_at: initiatedAt
        }));

        await this.run(() => addDoc(this.getBatchMessagesRef(this.chatBatch!.id), {
          sender_id: this.userId,
          sender_email: this.userName,
          sender_role: this.userRole,
          message: 'Procurement requested delivery (P.O uploaded). Production must accept to confirm batch delivery.',
          message_type: 'close_request',
          created_at: initiatedAt
        }));

        this.patchLocalBatch(this.chatBatch.id, {
          close_request_pending: true,
          close_request_production_accepted: false,
          close_request_initiated_at: initiatedAt,
          close_request_initiated_by: this.userId,
          close_request_mode: 'procurement'
        });
        this.showToast('Batch delivery request initiated', 'success');
      } catch (err) {
        console.error('initiateCloseRequestInChat batch failed', err);
        this.showToast('Failed to initiate delivery request', 'error');
      }
      return;
    }

    if (!this.chatTable) return;

    const mode: 'procurement' | 'production' =
      this.userRole === 'procurement' ||
      (this.userRole === 'admin' && this.isProcurementInWorkflow(this.chatTable))
        ? 'procurement'
        : 'production';
    const confirmMessage = mode === 'procurement'
      ? 'Mark request as delivered? User and Production must both accept to confirm delivery.'
      : 'Mark request as delivered? User must accept to confirm delivery.';
    if (!confirm(confirmMessage)) return;

    const tableRef = doc(this.firestore, 'tables', this.chatTable.id);
    const initiatedAt = new Date().toISOString();
    const systemMessage = mode === 'procurement'
      ? 'Procurement requested delivery (P.O uploaded). User and Production must accept to confirm all items as delivered.'
      : 'Production requested delivery. User must accept to confirm all items as delivered.';
    try {
      await this.run(() => updateDoc(tableRef, {
        close_request_pending: true,
        close_request_initiated_at: initiatedAt,
        close_request_initiated_by: this.userId,
        close_request_mode: mode,
        close_request_user_accepted: false,
        close_request_production_accepted: false,
        updated_at: initiatedAt
      }));

      this.patchLocalTable(this.chatTable.id, {
        close_request_pending: true,
        close_request_user_accepted: false,
        close_request_production_accepted: false,
        close_request_initiated_at: initiatedAt,
        close_request_initiated_by: this.userId,
        close_request_mode: mode
      });

      const messagesRef = this.getTableMessagesRef(this.chatTable.id);
      try {
        await this.run(() => addDoc(messagesRef, {
          sender_id: this.userId,
          sender_email: this.userName,
          sender_role: this.userRole,
          message: systemMessage,
          message_type: 'close_request',
          created_at: initiatedAt
        }));
      } catch (msgErr) {
        console.error('initiateCloseRequestInChat message failed', msgErr);
        this.showToast('Delivery request saved, but chat message could not be posted', 'info');
        return;
      }

      this.showToast('Delivery request initiated', 'success');
    } catch (err) {
      console.error('initiateCloseRequestInChat failed', err);
      this.showToast(this.formatFirestoreError(err, 'Failed to initiate delivery request'), 'error');
    }
  }

  async respondToCloseRequest(accept: boolean) {
    if (this.chatContextType === 'batch' && this.chatBatch?.close_request_pending) {
      if (!this.canRespondToCloseRequest()) return;
      try {
        const batchId = this.chatBatch.id;
        const batchRef = doc(this.firestore, 'procurement_batches', batchId);
        const now = new Date().toISOString();

        if (!accept) {
          await this.run(() => updateDoc(batchRef, {
            close_request_pending: false,
            close_request_production_accepted: false,
            close_request_mode: null,
            updated_at: now
          }));
          await this.addBatchChatSystemMessage(
            batchId,
            'Production rejected the batch delivery request. Batch remains PENDING.',
            'close_response',
            'rejected'
          );
          this.patchLocalBatch(batchId, {
            close_request_pending: false,
            close_request_production_accepted: false,
            close_request_mode: undefined
          });
          this.showToast('Batch delivery request rejected', 'info');
          return;
        }

        await this.run(() => updateDoc(batchRef, {
          close_request_production_accepted: true,
          updated_at: now
        }));
        await this.addBatchChatSystemMessage(
          batchId,
          'Production accepted the batch delivery request.',
          'close_response',
          'accepted'
        );

        const batchSnap = await this.run(() => getDoc(batchRef));
        const data = batchSnap.data() || {};
        if (data['close_request_production_accepted'] === true) {
          await this.finalizeBatchCloseRequest(batchId);
        } else {
          this.patchLocalBatch(batchId, { close_request_production_accepted: true });
          this.showToast('Your acceptance was recorded', 'success');
        }
      } catch {
        this.showToast('Failed to respond to batch delivery request', 'error');
      }
      return;
    }

    if (!this.chatTable?.close_request_pending || !this.canRespondToCloseRequest()) return;

    try {
      const tableId = this.chatTable.id;
      const tableRef = doc(this.firestore, 'tables', tableId);
      const now = new Date().toISOString();
      const roleLabel = this.userRole === 'production' ? 'Production' : 'User';

      if (!accept) {
        await this.run(() => updateDoc(tableRef, {
          close_request_pending: false,
          close_request_user_accepted: false,
          close_request_production_accepted: false,
          close_request_mode: null,
          updated_at: now
        }));
        await this.addChatSystemMessage(
          tableId,
          `${roleLabel} rejected the delivery request. Request remains PENDING.`,
          'close_response',
          'rejected'
        );
        this.patchLocalTable(tableId, {
          close_request_pending: false,
          close_request_user_accepted: false,
          close_request_production_accepted: false,
          close_request_mode: undefined
        });
        this.showToast('Delivery request rejected', 'info');
        return;
      }

      const updates = {
        updated_at: now,
        ...(this.userRole === 'user' || this.userRole === 'store'
          ? { close_request_user_accepted: true as const }
          : this.userRole === 'production'
            ? { close_request_production_accepted: true as const }
            : {})
      };

      await this.run(() => updateDoc(tableRef, updates));
      await this.addChatSystemMessage(
        tableId,
        `${roleLabel} accepted the delivery request.`,
        'close_response',
        'accepted'
      );

      const tableSnap = await this.run(() => getDoc(tableRef));
      const data = tableSnap.data() || {};
      const userAccepted = data['close_request_user_accepted'] === true;
      const productionAccepted = data['close_request_production_accepted'] === true;
      const mode: 'procurement' | 'production' =
        data['close_request_mode'] === 'production' ? 'production' : 'procurement';

      this.patchLocalTable(tableId, {
        close_request_user_accepted: userAccepted,
        close_request_production_accepted: productionAccepted
      });

      const deliveryConfirmed = mode === 'production'
        ? userAccepted
        : userAccepted && productionAccepted;

      if (deliveryConfirmed) {
        await this.finalizeCloseRequest(tableId);
      } else {
        this.showToast('Your acceptance was recorded', 'success');
      }
    } catch {
      this.showToast('Failed to respond to delivery request', 'error');
    }
  }

  private async finalizeCloseRequest(tableId: string) {
    const now = new Date().toISOString();
    const tableRef = doc(this.firestore, 'tables', tableId);

    await this.markTableRequisitionsDelivered(tableId);

    await this.run(() => updateDoc(tableRef, {
      request_status: 'DONE',
      request_closed: true,
      request_closed_at: now,
      request_closed_by: this.userId,
      close_request_pending: false,
      close_request_mode: null,
      updated_at: now
    }));
    await this.addChatSystemMessage(
      tableId,
      'Request delivered. All items marked as delivered. Chat is now read-only.',
      'system'
    );
    this.patchLocalTable(tableId, {
      request_status: 'DONE',
      request_closed: true,
      request_closed_at: now,
      request_closed_by: this.userId,
      close_request_pending: false,
      close_request_mode: undefined
    });
    this.showToast('Request delivered', 'success');
  }

  private async finalizeBatchCloseRequest(batchId: string) {
    const now = new Date().toISOString();
    const batch = this.getBatchById(batchId);
    if (!batch) return;

    const batchRef = doc(this.firestore, 'procurement_batches', batchId);
    await this.run(() => updateDoc(batchRef, {
      request_status: 'DONE',
      request_closed: true,
      request_closed_at: now,
      request_closed_by: this.userId,
      close_request_pending: false,
      close_request_mode: null,
      updated_at: now
    }));

    await this.addBatchChatSystemMessage(
      batchId,
      'Batch delivered. Procurement chat is now read-only. Production will coordinate delivery with each user separately.',
      'system'
    );

    for (const tableId of batch.table_ids) {
      const tableRef = doc(this.firestore, 'tables', tableId);
      await this.run(() => updateDoc(tableRef, {
        procurement_batch_closed: true,
        user_delivery_chat_started_at: now,
        close_request_pending: false,
        close_request_user_accepted: false,
        close_request_production_accepted: false,
        close_request_mode: null,
        updated_at: now
      }));
      await this.addChatSystemMessage(
        tableId,
        'You have been added to this delivery chat. Production will coordinate S.M upload and delivery confirmation with you.',
        'system'
      );
      this.patchLocalTable(tableId, {
        procurement_batch_closed: true,
        user_delivery_chat_started_at: now,
        close_request_pending: false,
        close_request_user_accepted: false,
        close_request_production_accepted: false,
        close_request_mode: undefined
      });
    }

    this.patchLocalBatch(batchId, {
      request_status: 'DONE',
      request_closed: true,
      request_closed_at: now,
      request_closed_by: this.userId,
      close_request_pending: false,
      close_request_mode: undefined
    });
    this.setupUnreadChatListeners();
    this.showToast('Batch delivered — opening user delivery chat', 'success');
    this.autoOpenUserDeliveryChat(batchId);
  }

  private async addBatchChatSystemMessage(
    batchId: string,
    message: string,
    messageType: ChatMessage['message_type'],
    closeResponse?: 'accepted' | 'rejected'
  ) {
    const messagesRef = this.getBatchMessagesRef(batchId);
    const payload: Record<string, unknown> = {
      sender_id: this.userId,
      sender_email: this.userName,
      sender_role: this.userRole,
      message,
      message_type: messageType,
      created_at: new Date().toISOString()
    };
    if (closeResponse) payload['close_response'] = closeResponse;
    await this.run(() => addDoc(messagesRef, payload));
  }

  private async markTableRequisitionsDelivered(tableId: string) {
    const reqs = [
      ...this.requisitions,
      ...this.productionSubmissions,
      ...this.productionReviewed,
      ...this.procurementReviewed
    ].filter(r => r.table_id === tableId);

    const seen = new Set<string>();
    for (const req of reqs) {
      if (!req.id || seen.has(req.id)) continue;
      seen.add(req.id);
      if (req.status === 'Delivered' || req.status === 'Removed') continue;
      const success = await this.db.updateRequisitionStatus(
        req.id,
        'Delivered',
        this.userId,
        tableId,
        {}
      );
      if (success) {
        req.status = 'Delivered';
      }
    }
  }

  private async addChatSystemMessage(
    tableId: string,
    message: string,
    messageType: ChatMessage['message_type'],
    closeResponse?: 'accepted' | 'rejected'
  ) {
    const messagesRef = this.getTableMessagesRef(tableId);
    const payload: Record<string, unknown> = {
      sender_id: this.userId,
      sender_email: this.userName,
      sender_role: this.userRole,
      message,
      message_type: messageType,
      created_at: new Date().toISOString()
    };
    if (closeResponse) payload['close_response'] = closeResponse;
    await this.run(() => addDoc(messagesRef, payload));
  }

  async onChatSmFileSelected(event: Event) {
    if (!this.chatTable || !this.isChatWritable() || !this.canUploadSm()) {
      this.showToast('Only production can upload Shipping Manifest files', 'error');
      return;
    }
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    if (file.size > this.maxPoFileBytes) {
      this.showToast('S.M. file must be 500KB or smaller', 'error');
      input.value = '';
      return;
    }

    try {
      this.isUploadingSm = true;
      const { base64, mime } = await this.readFileAsBase64(file);
      const tableRef = doc(this.firestore, 'tables', this.chatTable.id);
      const now = new Date().toISOString();

      console.log('SM upload attempt', {
        tableId: this.chatTable.id,
        userId: this.userId,
        userRole: this.userRole,
        submitted: this.chatTable.submitted,
        production_reviewed: this.chatTable.production_reviewed,
        procurement_batch_closed: this.chatTable.procurement_batch_closed
      });

      await this.run(() => updateDoc(tableRef, {
        sm_file_data: base64,
        sm_file_mime: mime,
        sm_file_name: file.name,
        sm_file_size: file.size,
        sm_file_type: mime,
        sm_uploaded_at: now,
        sm_uploaded_by: this.userId,
        updated_at: now
      }));

      this.patchLocalTable(this.chatTable.id, {
        sm_file_data: base64,
        sm_file_mime: mime,
        sm_file_name: file.name,
        sm_file_size: file.size,
        sm_file_type: mime
      });

      const messagesRef = this.getTableMessagesRef(this.chatTable.id);
      try {
        await this.run(() => addDoc(messagesRef, {
          sender_id: this.userId,
          sender_email: this.userName,
          sender_role: this.userRole,
          message: `Uploaded Shipping Manifest: ${file.name}`,
          message_type: 'sm_upload',
          created_at: now
        }));
      } catch (msgErr) {
        console.error('onChatSmFileSelected message failed', msgErr);
        this.showToast('Shipping Manifest saved, but chat message could not be posted', 'info');
        return;
      }

      this.showToast('Shipping Manifest uploaded', 'success');
    } catch (err) {
      console.error('onChatSmFileSelected failed', err);
      this.showToast(this.formatFirestoreError(err, 'Failed to upload Shipping Manifest'), 'error');
    } finally {
      this.isUploadingSm = false;
      input.value = '';
    }
  }

  async onChatImageSelected(event: Event) {
    if (!this.isChatWritable()) return;
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    if (!file.type.startsWith('image/')) {
      this.showToast('Please select an image file', 'error');
      input.value = '';
      return;
    }

    if (file.size > this.maxChatImageBytes) {
      this.showToast('Image must be 500KB or smaller', 'error');
      input.value = '';
      return;
    }

    try {
      const { base64, mime } = await this.readFileAsBase64(file);
      const messagesRef = this.chatContextType === 'batch' && this.chatBatch
        ? this.getBatchMessagesRef(this.chatBatch.id)
        : this.chatTable
          ? this.getTableMessagesRef(this.chatTable.id)
          : null;
      if (!messagesRef) return;
      const now = new Date().toISOString();
      await this.run(() => addDoc(messagesRef, {
        sender_id: this.userId,
        sender_email: this.userName,
        sender_role: this.userRole,
        message: '',
        message_type: 'image',
        image_data: base64,
        image_mime: mime,
        created_at: now
      }));
    } catch (err) {
      console.error('onChatImageSelected failed', err);
      this.showToast('Failed to send image', 'error');
    } finally {
      input.value = '';
    }
  }

  async onChatPoFileSelected(event: Event) {
    if (!this.isChatWritable() || !this.canUploadPo()) {
      this.showToast('Only procurement can upload P.O files', 'error');
      return;
    }
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    if (file.size > this.maxPoFileBytes) {
      this.showToast('P.O file must be 500KB or smaller', 'error');
      input.value = '';
      return;
    }

    try {
      this.isUploadingPo = true;
      const { base64, mime } = await this.readFileAsBase64(file);
      const now = new Date().toISOString();

      if (this.chatContextType === 'batch' && this.chatBatch) {
        const batchRef = doc(this.firestore, 'procurement_batches', this.chatBatch.id);
        await this.run(() => updateDoc(batchRef, {
          po_file_data: base64,
          po_file_mime: mime,
          po_file_name: file.name,
          po_file_size: file.size,
          po_file_type: mime,
          po_uploaded_at: now,
          po_uploaded_by: this.userId,
          updated_at: now
        }));

        await this.run(() => addDoc(this.getBatchMessagesRef(this.chatBatch!.id), {
          sender_id: this.userId,
          sender_email: this.userName,
          sender_role: this.userRole,
          message: `Uploaded P.O: ${file.name}`,
          message_type: 'po_upload',
          po_file_name: file.name,
          created_at: now
        }));

        this.patchLocalBatch(this.chatBatch.id, {
          po_file_data: base64,
          po_file_mime: mime,
          po_file_name: file.name,
          po_file_size: file.size,
          po_file_type: mime
        });
      } else if (this.chatTable) {
        const tableRef = doc(this.firestore, 'tables', this.chatTable.id);
        await this.run(() => updateDoc(tableRef, {
          po_file_data: base64,
          po_file_mime: mime,
          po_file_name: file.name,
          po_file_size: file.size,
          po_file_type: mime,
          po_uploaded_at: now,
          po_uploaded_by: this.userId,
          updated_at: now
        }));

        const messagesRef = this.getTableMessagesRef(this.chatTable.id);
        await this.run(() => addDoc(messagesRef, {
          sender_id: this.userId,
          sender_email: this.userName,
          sender_role: this.userRole,
          message: `Uploaded P.O: ${file.name}`,
          message_type: 'po_upload',
          po_file_name: file.name,
          created_at: now
        }));

        this.patchLocalTable(this.chatTable.id, {
          po_file_data: base64,
          po_file_mime: mime,
          po_file_name: file.name,
          po_file_size: file.size,
          po_file_type: mime
        });
      }
      this.showToast('P.O uploaded to chat', 'success');
    } catch {
      this.showToast('Failed to upload P.O', 'error');
    } finally {
      this.isUploadingPo = false;
      input.value = '';
    }
  }

  private patchLocalBatch(batchId: string, patch: Partial<ProcurementBatch>) {
    const batchIndex = this.procurementBatches.findIndex(b => b.id === batchId);
    if (batchIndex !== -1) {
      this.procurementBatches[batchIndex] = { ...this.procurementBatches[batchIndex], ...patch };
    }
    if (this.selectedBatch?.id === batchId) {
      this.selectedBatch = { ...this.selectedBatch, ...patch };
    }
    if (this.chatBatch?.id === batchId) {
      this.chatBatch = { ...this.chatBatch, ...patch };
    }
  }

  private patchLocalTable(tableId: string, patch: Partial<Table>) {
    const sanitizedPatch = (this.userRole === 'user' || this.userRole === 'store')
      ? Object.fromEntries(Object.entries(patch).filter(([key]) => !key.startsWith('po_')))
      : patch;

    const tableIndex = this.tables.findIndex(t => t.id === tableId);
    if (tableIndex !== -1) {
      this.tables[tableIndex] = this.sanitizeTableForRole({ ...this.tables[tableIndex], ...sanitizedPatch });
    }
    if (this.selectedTable?.id === tableId) {
      this.selectedTable = this.sanitizeTableForRole({ ...this.selectedTable, ...sanitizedPatch });
    }
    if (this.chatTable?.id === tableId) {
      this.chatTable = this.sanitizeTableForRole({ ...this.chatTable, ...sanitizedPatch });
    }
  }

  ngOnDestroy() {
    if (this.chatUnsubscribe) {
      this.chatUnsubscribe();
      this.chatUnsubscribe = null;
    }
    this.teardownUnreadChatListeners();
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
    this.updateProcurementPagination();
  }

  updateProcurementPagination() {
    const materialCount = this.userRole === 'procurement' && this.selectedBatch
      ? this.procurementConsolidatedMaterials.length
      : this.procurementTableSummaries.length;
    this.procurementTotalPages = Math.max(1, Math.ceil(materialCount / this.pageSize));
    if (this.currentPage > this.procurementTotalPages) {
      this.currentPage = this.procurementTotalPages;
    }
    const start = (this.currentPage - 1) * this.pageSize;
    if (this.userRole === 'procurement' && this.selectedBatch) {
      this.paginatedProcurementMaterials = this.procurementConsolidatedMaterials.slice(start, start + this.pageSize);
      this.paginatedProcurementSummaries = [];
    } else {
      this.paginatedProcurementMaterials = [];
      this.paginatedProcurementSummaries = this.procurementTableSummaries.slice(start, start + this.pageSize);
    }
  }

  goToPage(page: number) {
    if (this.userRole === 'procurement') {
      if (page >= 1 && page <= this.procurementTotalPages) {
        this.currentPage = page;
        this.updateProcurementPagination();
      }
      return;
    }
    if (page >= 1 && page <= this.totalPages) {
      this.currentPage = page;
      this.updatePagination();
    }
  }

  onPageSizeChange() {
    this.currentPage = 1;
    this.updatePagination();
    this.updateProcurementPagination();
  }

  private formatFirestoreError(err: unknown, fallback: string): string {
    const code = (err as { code?: string })?.code || '';
    const message = (err as { message?: string })?.message || '';

    if (code === 'permission-denied') {
      return 'Permission denied. Deploy the latest firestore.rules to Firebase, then try again.';
    }
    if (code === 'invalid-argument' || message.toLowerCase().includes('size')) {
      return 'File is too large to store on this table. Try a smaller file.';
    }
    return fallback;
  }

  showToast(message: string, type: 'success' | 'error' | 'info' = 'info') {
    this.toast.show(message, type);
  }

  hideSnackbar() {
    this.toast.dismissAll();
  }
}