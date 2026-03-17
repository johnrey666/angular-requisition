import { Component, OnInit, HostListener } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import * as ExcelJS from 'exceljs';
import { DatabaseService } from '../../../core/services/database.service';
import { AuthService } from '../../../core/services/auth.service';
import { LoaderService } from '../../../core/services/loader.service';
import { Firestore, doc, getDoc, collection, query, where, getDocs } from '@angular/fire/firestore';
import { Router } from '@angular/router';

interface Material {
  raw_material: string;
  quantity_per_batch: number | null;
  unit: string;
  type: string;
}

interface ProductionItem {
  id: string;
  sku_code: string;
  sku_name: string;
  category: string;
  supplier: string;
  qty_needed: number;
  table_id: string;
  user_id: string;
  materials?: Material[];
  materialCount?: number;
  totalRequired?: number;
  created_at?: string;
}

interface ProductionTable {
  id: string;
  name: string;
  user_id: string;
  type: 'production' | 'inventory';
  item_count?: number;
  created_at?: string;
  updated_at?: string;
}

@Component({
  selector: 'app-page2',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './page2.component.html',
  styleUrls: ['./page2.component.css']
})
export class Page2Component implements OnInit {
  newItem = {
    category: '',
    sku_code: '',
    sku_name: '',
    supplier: '',
    qty_needed: null as number | null
  };

  categories: string[] = [];
  availableSkus: { sku_code: string; sku_name: string }[] = [];
  productionItems: ProductionItem[] = [];
  filteredItems: ProductionItem[] = [];
  paginatedItems: ProductionItem[] = [];
  
  // Table Management
  productionTables: ProductionTable[] = [];
  currentTable: ProductionTable | null = null;
  showTableDropdown = false;

  uploadingMaster = false;
  uploadStatus = '';
  addingItem = false;
  expandedRows: { [id: string]: boolean } = {};
  loadingMaterials: { [id: string]: boolean } = {};

  searchQuery = '';
  filterCategory = '';
  
  // Pagination
  currentPage = 1;
  pageSize = 10;
  totalPages = 1;
  
  // Snackbar
  showSnackbar = false;
  snackbarMessage = '';
  snackbarType: 'success' | 'error' | 'info' = 'info';
  snackbarTimeout: any;

  // User Role
  userRole: string = '';
  userId: string = '';

  // Allowed roles for accessing this page
  private readonly allowedRoles = ['user', 'store', 'production', 'procurement', 'admin'];

  // Expose Math to template
  Math = Math;

  constructor(
    private db: DatabaseService,
    private auth: AuthService,
    private firestore: Firestore,
    private router: Router,
    private loader: LoaderService
  ) {}

  async ngOnInit() {
    const user = await this.auth.getCurrentUserPromise();
    if (user) {
      this.userId = user.uid;
      await this.loadUserRole();
      await this.loadCategories(); // Load categories first
      await this.loadProductionTablesDirectly();
    } else {
      this.showToast('Please log in to continue', 'error');
      this.router.navigate(['/login']);
    }
  }

  async loadUserRole() {
    try {
      const userDoc = await getDoc(doc(this.firestore, 'users', this.userId));
      if (userDoc.exists()) {
        const data = userDoc.data() as any;
        this.userRole = data['role'] || 'user';
        console.log('User role loaded:', this.userRole);
        
        if (!this.allowedRoles.includes(this.userRole)) {
          this.showToast('You do not have access to Production Management', 'error');
          this.router.navigate(['/dashboard']);
        }
      } else {
        this.userRole = 'user';
        console.log('No role document found, defaulting to user');
      }
    } catch (err) {
      console.error('Failed to load user role', err);
      this.userRole = 'user';
    }
  }

  // Direct Firestore query to load production tables
  async loadProductionTablesDirectly() {
    try {
      this.loader.show('Loading production lines...');
      if (!this.userId) {
        this.showToast('You must be logged in', 'error');
        this.loader.hide();
        return;
      }
      
      console.log('Loading production tables directly for user:', this.userId);
      
      const tablesRef = collection(this.firestore, 'tables');
      const q = query(
        tablesRef, 
        where('user_id', '==', this.userId),
        where('type', '==', 'production')
      );
      
      const querySnapshot = await getDocs(q);
      console.log('Found', querySnapshot.size, 'production tables');
      
      const loadedTables: ProductionTable[] = [];
      querySnapshot.forEach(doc => {
        const data = doc.data();
        console.log('Production table document:', { id: doc.id, ...data });
        
        loadedTables.push({
          id: doc.id,
          name: data['name'] || 'Untitled',
          user_id: data['user_id'] || this.userId,
          type: data['type'] || 'production',
          item_count: data['item_count'] || 0,
          created_at: data['created_at'],
          updated_at: data['updated_at']
        });
      });
      
      console.log('Loaded production tables:', loadedTables);
      this.productionTables = loadedTables;
      
      const lastTableId = localStorage.getItem(`lastSelectedProductionTable_${this.userId}`);
      if (lastTableId && this.productionTables.some(t => t.id === lastTableId)) {
        this.currentTable = this.productionTables.find(t => t.id === lastTableId) || null;
        console.log('Restored last selected production table:', this.currentTable);
      } else if (this.productionTables.length > 0) {
        this.currentTable = this.productionTables[0];
        console.log('Defaulting to first production table:', this.currentTable);
      }
      
      if (this.currentTable) {
        await this.loadProductionItemsDirectly();
      } else {
        this.productionItems = [];
        this.filteredItems = [];
        this.updatePagination();
      }
    } catch (err) {
      console.error('Failed to load production tables directly:', err);
      this.productionTables = [];
      this.currentTable = null;
      this.showToast('Failed to load production tables', 'error');
    } finally {
      this.loader.hide();
    }
  }

  // Direct Firestore query to load production items
  async loadProductionItemsDirectly() {
    if (!this.currentTable || !this.userId) {
      console.log('Missing table or userId for loading production items');
      return;
    }

    try {
      this.loader.show('Loading items...');
      console.log('Loading production items for table:', this.currentTable.id);
      
      const requisitionsRef = collection(this.firestore, 'requisitions');
      const q = query(
        requisitionsRef,
        where('table_id', '==', this.currentTable.id),
        where('user_id', '==', this.userId)
      );
      
      const querySnapshot = await getDocs(q);
      console.log('Found', querySnapshot.size, 'production items');
      
      const loadedItems: ProductionItem[] = [];
      querySnapshot.forEach(doc => {
        const data = doc.data();
        loadedItems.push({
          id: doc.id,
          sku_code: data['sku_code'] || '',
          sku_name: data['sku_name'] || '',
          category: data['category'] || '',
          supplier: data['supplier'] || '',
          qty_needed: data['qty_needed'] || data['quantity'] || 0,
          table_id: data['table_id'] || this.currentTable!.id,
          user_id: data['user_id'] || this.userId,
          created_at: data['created_at'],
          materialCount: 0,
          totalRequired: 0
        });
      });
      
      console.log('Loaded production items:', loadedItems);
      this.productionItems = loadedItems;
      this.applyFilter();
      
    } catch (err) {
      console.error('Failed to load production items:', err);
      this.productionItems = [];
      this.filteredItems = [];
      this.updatePagination();
      this.showToast('Failed to load production items', 'error');
    } finally {
      this.loader.hide();
    }
  }

  async loadCategories() {
    try {
      console.log('Loading categories from master data...');
      this.categories = await this.db.getUniqueCategories();
      console.log('Loaded categories:', this.categories);
      
      // If no categories found, show message
      if (this.categories.length === 0) {
        console.log('No categories found in master data');
        this.uploadStatus = 'No categories found. Please upload master data first.';
      }
    } catch (err) {
      console.error('Failed to load categories', err);
      this.categories = [];
      this.showToast('Failed to load categories', 'error');
    }
  }

  async loadProductionTables() {
    try {
      if (!this.userId) {
        this.showToast('You must be logged in', 'error');
        return;
      }
      
      this.productionTables = await this.db.getUserTablesByType(this.userId, 'production');
      console.log('Loaded production tables:', this.productionTables);
      
      const lastTableId = localStorage.getItem(`lastSelectedProductionTable_${this.userId}`);
      if (lastTableId && this.productionTables.some(t => t.id === lastTableId)) {
        this.currentTable = this.productionTables.find(t => t.id === lastTableId) || null;
      } else if (this.productionTables.length > 0) {
        this.currentTable = this.productionTables[0];
      }
      
      if (this.currentTable) {
        await this.loadProductionItems();
      } else {
        this.productionItems = [];
        this.filteredItems = [];
        this.updatePagination();
      }
    } catch (err) {
      console.error('Failed to load production tables', err);
      this.productionTables = [];
      this.currentTable = null;
      this.showToast('Failed to load tables', 'error');
    }
  }

  async loadProductionItems() {
    if (!this.currentTable) {
      this.productionItems = [];
      this.filteredItems = [];
      this.updatePagination();
      return;
    }

    try {
      if (!this.userId) {
        this.showToast('You must be logged in', 'error');
        return;
      }

      const requisitions = await this.db.getTableRequisitions(this.currentTable.id, this.userId);
      console.log('Loaded requisitions:', requisitions);
      
      this.productionItems = requisitions.map((req: any) => ({
        id: req.id,
        sku_code: req.sku_code,
        sku_name: req.sku_name,
        category: req.category,
        supplier: req.supplier,
        qty_needed: req.qty_needed || req.quantity || 0,
        table_id: req.table_id,
        user_id: req.user_id,
        created_at: req.created_at,
        materialCount: 0,
        totalRequired: 0
      }));
      
      this.applyFilter();
    } catch (err) {
      console.error('Failed to load production items', err);
      this.productionItems = [];
      this.filteredItems = [];
      this.updatePagination();
      this.showToast('Failed to load production items', 'error');
    }
  }

  async onCategoryChange() {
    this.newItem.sku_code = '';
    this.newItem.sku_name = '';

    if (!this.newItem.category) {
      this.availableSkus = [];
      return;
    }

    try {
      console.log('Loading SKUs for category:', this.newItem.category);
      this.availableSkus = await this.db.getSkusByCategory(this.newItem.category);
      console.log('Available SKUs for category:', this.availableSkus);
      
      if (this.availableSkus.length === 0) {
        this.showToast('No SKUs found for this category', 'info');
      }
    } catch (err) {
      console.error('Failed to load SKUs', err);
      this.availableSkus = [];
      this.showToast('Failed to load SKUs for selected category', 'error');
    }
  }

  onSkuChange() {
    const found = this.availableSkus.find(s => s.sku_code === this.newItem.sku_code);
    this.newItem.sku_name = found?.sku_name || '';
  }

  async onFileSelected(event: Event) {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    // Check if user is admin
    if (this.userRole !== 'admin') {
      this.showToast('Only administrators can upload master data', 'error');
      input.value = '';
      return;
    }

    this.uploadingMaster = true;
    this.uploadStatus = 'Uploading master data...';

    try {
      const result = await this.db.uploadMasterData(file);
      if (result.success) {
        this.uploadStatus = `Successfully imported ${result.count || 0} rows`;
        await this.loadCategories(); // Reload categories after upload
        this.showToast(`Successfully imported ${result.count || 0} rows`, 'success');
      } else {
        this.uploadStatus = `Upload failed: ${result.error || 'Unknown error'}`;
        this.showToast('Upload failed', 'error');
      }
    } catch (err: any) {
      this.uploadStatus = `Error: ${err.message || 'Failed to process file'}`;
      this.showToast('Error uploading file', 'error');
    } finally {
      this.uploadingMaster = false;
      input.value = '';
    }
  }

  canAddItem(): boolean {
    return !!this.currentTable &&
           !!this.newItem.category &&
           !!this.newItem.sku_code &&
           !!this.newItem.supplier?.trim() &&
           this.newItem.qty_needed != null && this.newItem.qty_needed > 0;
  }

  async addItem() {
    if (!this.canAddItem()) {
      this.showToast('Please fill all required fields and select a table', 'error');
      return;
    }

    this.addingItem = true;
    this.loader.show('Adding item...');
    
    if (!this.userId) {
      this.showToast('You must be logged in', 'error');
      this.addingItem = false;
      return;
    }

    if (!this.currentTable) {
      this.showToast('No table selected', 'error');
      this.addingItem = false;
      return;
    }

    const entry = {
      sku_code: this.newItem.sku_code,
      sku_name: this.newItem.sku_name,
      category: this.newItem.category,
      supplier: this.newItem.supplier.trim(),
      qty_needed: this.newItem.qty_needed!,
      table_id: this.currentTable.id,
      user_id: this.userId
    };

    try {
      const result = await this.db.createRequisition(entry, []);
      
      if (result.success && result.id) {
        const newItem: ProductionItem = { 
          id: result.id, 
          ...entry, 
          materialCount: 0,
          totalRequired: 0
        };
        
        this.productionItems.unshift(newItem);
        
        const newCount = (this.currentTable.item_count || 0) + 1;
        await this.db.updateTableItemCount(this.currentTable.id, newCount, this.userId);
        this.currentTable.item_count = newCount;
        
        await this.loadProductionTablesDirectly();
        
        this.newItem = { category: '', sku_code: '', sku_name: '', supplier: '', qty_needed: null };
        this.availableSkus = [];
        
        this.applyFilter();
        
        this.showToast('Item added to production successfully', 'success');
      } else {
        this.showToast('Failed to save item', 'error');
      }
    } catch (err) {
      console.error('Add item error', err);
      this.showToast('Error adding item', 'error');
    } finally {
      this.addingItem = false;
      this.loader.hide();
    }
  }

  async toggleRow(item: ProductionItem) {
    if (!item.id) return;
    
    this.expandedRows[item.id] = !this.expandedRows[item.id];

    if (this.expandedRows[item.id] && !item.materials) {
      this.loadingMaterials[item.id] = true;
      try {
        const materials = await this.db.getMaterialsForSku(item.sku_code);
        item.materials = materials || [];
        item.materialCount = item.materials?.length || 0;
      } catch (err) {
        console.error('Failed to load materials', err);
        item.materials = [];
        item.materialCount = 0;
        this.showToast('Failed to load materials', 'error');
      } finally {
        this.loadingMaterials[item.id] = false;
      }
    }
  }

  applyFilter() {
    let list = [...this.productionItems];

    if (this.filterCategory) {
      list = list.filter(i => i.category === this.filterCategory);
    }

    if (this.searchQuery.trim()) {
      const q = this.searchQuery.toLowerCase();
      list = list.filter(i =>
        i.sku_code.toLowerCase().includes(q) ||
        (i.sku_name || '').toLowerCase().includes(q) ||
        (i.supplier || '').toLowerCase().includes(q)
      );
    }

    this.filteredItems = list;
    this.currentPage = 1;
    this.updatePagination();
  }

  clearFilters() {
    this.searchQuery = '';
    this.filterCategory = '';
    this.applyFilter();
    this.showToast('Filters cleared', 'info');
  }

  get totalQuantity(): number {
    return this.filteredItems.reduce((sum, i) => sum + (i.qty_needed || 0), 0);
  }

  calculateMaterialTotal(itemQty: number | null, qtyPerBatch: number | null): number {
    const qty = itemQty || 0;
    const batchQty = qtyPerBatch || 0;
    return batchQty * qty;
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

  async selectTable(table: ProductionTable) {
    if (table.type !== 'production') {
      this.showToast('Invalid table type', 'error');
      return;
    }

    if (table.user_id !== this.userId) {
      this.showToast('You can only access your own tables', 'error');
      return;
    }

    this.currentTable = table;
    this.showTableDropdown = false;
    
    localStorage.setItem(`lastSelectedProductionTable_${this.userId}`, table.id);
    
    this.searchQuery = '';
    this.filterCategory = '';
    
    await this.loadProductionItemsDirectly();
    this.showToast(`Switched to production line: ${table.name}`, 'info');
  }

  async createNewTable() {
    const tableName = prompt('Enter production line name:');
    if (!tableName?.trim()) return;

    try {
      this.loader.show('Creating production line...');
      if (!this.userId) {
        this.showToast('You must be logged in to create tables', 'error');
        return;
      }

      const tableData = {
        user_id: this.userId,
        name: tableName.trim(),
        item_count: 0
      };

      const result = await this.db.createUserTable(tableData, 'production');
      if (result.success && result.tableId) {
        await this.loadProductionTablesDirectly();
        this.showToast(`Production line "${tableName}" created successfully`, 'success');
      } else {
        this.showToast('Failed to create production line', 'error');
      }
    } catch (err) {
      console.error('Failed to create table', err);
      this.showToast('Error creating production line', 'error');
    } finally {
      this.loader.hide();
    }
  }

  async renameTable(table: ProductionTable) {
    if (table.type !== 'production') {
      this.showToast('Invalid table type', 'error');
      return;
    }

    if (table.user_id !== this.userId) {
      this.showToast('You can only rename your own tables', 'error');
      return;
    }

    const newName = prompt('Enter new production line name:', table.name);
    if (!newName?.trim() || newName === table.name) return;

    try {
      if (!this.userId) {
        this.showToast('You must be logged in', 'error');
        return;
      }

      const success = await this.db.updateTableName(table.id, newName.trim(), this.userId);
      if (success) {
        table.name = newName.trim();
        await this.loadProductionTablesDirectly();
        this.showToast('Production line renamed successfully', 'success');
      } else {
        this.showToast('Failed to rename production line', 'error');
      }
    } catch (err) {
      console.error('Failed to rename table', err);
      this.showToast('Error renaming production line', 'error');
    }
  }

  async deleteTable(table: ProductionTable) {
    if (table.type !== 'production') {
      this.showToast('Invalid table type', 'error');
      return;
    }

    if (table.user_id !== this.userId) {
      this.showToast('You can only delete your own tables', 'error');
      return;
    }

    if (!confirm(`Are you sure you want to delete production line "${table.name}"? This will also delete all items in this line. This action cannot be undone.`)) {
      return;
    }

    try {
      if (!this.userId) {
        this.showToast('You must be logged in', 'error');
        return;
      }

      const success = await this.db.deleteTable(table.id, this.userId);
      if (success) {
        this.productionTables = this.productionTables.filter(t => t.id !== table.id);
        
        if (this.currentTable?.id === table.id) {
          if (this.productionTables.length > 0) {
            await this.selectTable(this.productionTables[0]);
          } else {
            this.currentTable = null;
            this.productionItems = [];
            this.filteredItems = [];
            this.updatePagination();
          }
        }
        
        this.showToast(`Production line "${table.name}" deleted successfully`, 'success');
      } else {
        this.showToast('Failed to delete production line', 'error');
      }
    } catch (err) {
      console.error('Failed to delete table', err);
      this.showToast('Error deleting production line', 'error');
    }
  }

  async deleteItem(item: ProductionItem, event: Event) {
    event.stopPropagation();
    
    if (!this.currentTable) {
      this.showToast('No production line selected', 'error');
      return;
    }

    if (confirm('Are you sure you want to delete this production item?')) {
      try {
        if (!this.userId) {
          this.showToast('You must be logged in', 'error');
          return;
        }

        const success = await this.db.deleteRequisition(item.id, this.userId, this.currentTable.id);
        
        if (success) {
          this.productionItems = this.productionItems.filter(i => i.id !== item.id);
          
          const newCount = Math.max(0, (this.currentTable.item_count || 0) - 1);
          await this.db.updateTableItemCount(this.currentTable.id, newCount, this.userId);
          this.currentTable.item_count = newCount;
          
          await this.loadProductionTablesDirectly();
          this.applyFilter();
          
          this.showToast('Production item deleted successfully', 'success');
        } else {
          this.showToast('Failed to delete production item', 'error');
        }
      } catch (err) {
        console.error('Failed to delete item', err);
        this.showToast('Error deleting production item', 'error');
      }
    }
  }

  async exportData() {
    if (!this.currentTable) {
      this.showToast('No production line selected', 'info');
      return;
    }

    if (this.filteredItems.length === 0) {
      this.showToast('No data to export', 'info');
      return;
    }

    try {
      this.showToast('Preparing production requirements export...', 'info');
      
      const fileName = `${this.currentTable.name.replace(/\s+/g, '_')}_production_requirements_${new Date().toISOString().split('T')[0]}.xlsx`;
      
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'Production Management System';
      workbook.lastModifiedBy = 'Production Management System';
      workbook.created = new Date();
      workbook.modified = new Date();
      
      const worksheet = workbook.addWorksheet('Production Requirements', {
        properties: {
          defaultColWidth: 15,
          showGridLines: true
        }
      });

      worksheet.mergeCells('A1:J1');
      const titleRow = worksheet.getRow(1);
      titleRow.getCell(1).value = `PRODUCTION REQUIREMENTS - ${this.currentTable.name}`;
      titleRow.getCell(1).font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
      titleRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C3E50' } };
      titleRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
      titleRow.height = 30;

      worksheet.mergeCells('A2:J2');
      const infoRow = worksheet.getRow(2);
      infoRow.getCell(1).value = `Production Manager: ${this.userRole} | Generated on: ${new Date().toLocaleString()} | Total Items: ${this.filteredItems.length}`;
      infoRow.getCell(1).font = { name: 'Arial', size: 11, italic: true };
      infoRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
      infoRow.height = 25;

      const headers = [
        'SKU Code',
        'Item Name',
        'Category',
        'Qty Needed',
        'Supplier',
        'Material',
        'Qty/Batch',
        'Unit',
        'Type',
        'Total Required'
      ];
      
      const headerRow = worksheet.addRow(headers);
      headerRow.eachCell((cell) => {
        cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2980B9' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
      headerRow.height = 30;

      for (const item of this.filteredItems) {
        if (!item.materials) {
          try {
            const materials = await this.db.getMaterialsForSku(item.sku_code);
            item.materials = materials || [];
            item.materialCount = item.materials?.length || 0;
          } catch (err) {
            console.error('Failed to load materials for export', err);
            item.materials = [];
          }
        }

        const materials = item.materials || [];

        if (materials.length === 0) {
          const row = worksheet.addRow([
            item.sku_code,
            item.sku_name || item.sku_code,
            item.category,
            item.qty_needed,
            item.supplier || '',
            'No materials',
            '',
            '',
            '',
            ''
          ]);
          this.styleDataRow(row);
          row.getCell(4).numFmt = '#,##0';
        } else {
          for (let i = 0; i < materials.length; i++) {
            const mat = materials[i];
            const row = worksheet.addRow([
              i === 0 ? item.sku_code : '',
              i === 0 ? (item.sku_name || item.sku_code) : '',
              i === 0 ? item.category : '',
              i === 0 ? item.qty_needed : '',
              i === 0 ? (item.supplier || '') : '',
              mat.raw_material || '',
              mat.quantity_per_batch || '',
              mat.unit || '',
              mat.type || '',
              this.calculateMaterialTotal(item.qty_needed, mat.quantity_per_batch)
            ]);
            this.styleDataRow(row);
            
            if (i === 0) {
              row.getCell(4).numFmt = '#,##0';
            }
            row.getCell(7).numFmt = '#,##0.00';
            row.getCell(10).numFmt = '#,##0.00';
          }
        }
        
        worksheet.addRow([]);
      }

      const widths = [18, 25, 15, 12, 20, 30, 15, 10, 15, 15];
      worksheet.columns.forEach((column, index) => {
        column.width = widths[index] || 15;
      });

      worksheet.views = [
        { state: 'frozen', xSplit: 0, ySplit: 3, activeCell: 'A4' }
      ];

      const buffer = await workbook.xlsx.writeBuffer();
      
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      
      link.setAttribute('href', url);
      link.setAttribute('download', fileName);
      link.style.visibility = 'hidden';
      
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      
      setTimeout(() => {
        URL.revokeObjectURL(url);
      }, 100);
      
      this.showToast(`Exported ${this.filteredItems.length} production items successfully`, 'success');
    } catch (err) {
      console.error('Export failed', err);
      this.showToast('Error exporting data: ' + (err instanceof Error ? err.message : 'Unknown error'), 'error');
    }
  }

  private styleDataRow(row: any) {
    row.eachCell((cell: any) => {
      cell.font = { name: 'Arial', size: 10 };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFBDC3C7' } },
        left: { style: 'thin', color: { argb: 'FFBDC3C7' } },
        bottom: { style: 'thin', color: { argb: 'FFBDC3C7' } },
        right: { style: 'thin', color: { argb: 'FFBDC3C7' } }
      };
    });
    
    [4, 7, 8, 9, 10].forEach(colIndex => {
      const cell = row.getCell(colIndex);
      if (cell.value) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    });
    
    [1, 2, 3, 5, 6].forEach(colIndex => {
      const cell = row.getCell(colIndex);
      if (cell.value) {
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      }
    });
  }

  showToast(message: string, type: 'success' | 'error' | 'info' = 'info') {
    this.snackbarMessage = message;
    this.snackbarType = type;
    this.showSnackbar = true;
    
    if (this.snackbarTimeout) {
      clearTimeout(this.snackbarTimeout);
    }
    
    this.snackbarTimeout = setTimeout(() => {
      this.hideSnackbar();
    }, 5000);
  }

  hideSnackbar() {
    this.showSnackbar = false;
    if (this.snackbarTimeout) {
      clearTimeout(this.snackbarTimeout);
      this.snackbarTimeout = null;
    }
  }

  updatePagination() {
    this.totalPages = Math.max(1, Math.ceil(this.filteredItems.length / this.pageSize));
    this.paginatedItems = this.filteredItems.slice(
      (this.currentPage - 1) * this.pageSize,
      this.currentPage * this.pageSize
    );
  }

  nextPage() {
    if (this.currentPage < this.totalPages) {
      this.currentPage++;
      this.updatePagination();
    }
  }

  prevPage() {
    if (this.currentPage > 1) {
      this.currentPage--;
      this.updatePagination();
    }
  }

  onPageSizeChange() {
    this.currentPage = 1;
    this.updatePagination();
  }

  getPageNumbers(): number[] {
    const pages: number[] = [];
    const total = this.totalPages;
    const current = this.currentPage;
    const delta = 2;
    
    if (total <= 7) {
      for (let i = 1; i <= total; i++) {
        pages.push(i);
      }
    } else {
      pages.push(1);
      
      let start = Math.max(2, current - delta);
      let end = Math.min(total - 1, current + delta);
      
      if (start > 2) {
        pages.push(-1);
      }
      
      for (let i = start; i <= end; i++) {
        pages.push(i);
      }
      
      if (end < total - 1) {
        pages.push(-1);
      }
      
      pages.push(total);
    }
    
    return pages;
  }

  goToPage(page: number) {
    if (page >= 1 && page <= this.totalPages && page !== this.currentPage) {
      this.currentPage = page;
      this.updatePagination();
    }
  }
}