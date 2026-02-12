import { Component, OnInit, HostListener } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DatabaseService } from '../../../core/services/database.service';
import { AuthService } from '../../../core/services/auth.service';

interface Material {
  raw_material: string;
  quantity_per_batch: number | null;
  unit: string;
  type: string;
}

interface InventoryItem {
  id: string;
  sku_code: string;
  sku_name: string;
  category: string;
  supplier: string;
  qty: number;
  table_id: string;
  user_id: string;
  materials?: Material[];
  materialCount?: number;
  totalRequired?: number;
}

interface UserTable {
  id: string;
  name: string;
  user_id: string;
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
    qty: null as number | null
  };

  categories: string[] = [];
  availableSkus: { sku_code: string; sku_name: string }[] = [];
  inventoryItems: InventoryItem[] = [];
  filteredItems: InventoryItem[] = [];
  paginatedItems: InventoryItem[] = [];
  
  // Table Management
  userTables: UserTable[] = [];
  currentTable: UserTable | null = null;
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

  // Expose Math to template
  Math = Math;

  constructor(
    private db: DatabaseService,
    private auth: AuthService
  ) {}

  async ngOnInit() {
    await this.loadCategories();
    await this.loadUserTables();
    // Don't load inventory until a table is selected
  }

  async loadCategories() {
    try {
      this.categories = await this.db.getUniqueCategories();
    } catch (err) {
      console.error('Failed to load categories', err);
      this.categories = [];
      this.showToast('Failed to load categories', 'error');
    }
  }

  async loadUserTables() {
    try {
      const userId = this.auth.getUserId();
      if (!userId) {
        this.showToast('You must be logged in', 'error');
        return;
      }
      
      this.userTables = await this.db.getUserTables(userId);
      
      // Don't create default table - user must create their own
      if (this.userTables.length === 0) {
        this.currentTable = null;
        this.inventoryItems = [];
        this.filteredItems = [];
        this.updatePagination();
      } else {
        // If we have tables but no current table selected, select first one
        if (!this.currentTable) {
          await this.selectTable(this.userTables[0]);
        }
      }
    } catch (err) {
      console.error('Failed to load user tables', err);
      this.userTables = [];
      this.currentTable = null;
      this.showToast('Failed to load tables', 'error');
    }
  }

  async loadInventory() {
    if (!this.currentTable) {
      this.inventoryItems = [];
      this.filteredItems = [];
      this.updatePagination();
      return;
    }

    try {
      const userId = this.auth.getUserId();
      if (!userId) {
        this.showToast('You must be logged in', 'error');
        return;
      }

      // Load only items belonging to the current table and current user
      const items = await this.db.getInventoryItemsByTable(this.currentTable.id, userId);
      this.inventoryItems = items.map((item: any) => ({
        ...item,
        materialCount: 0,
        totalRequired: 0
      }));
      this.applyFilter();
    } catch (err) {
      console.error('Failed to load inventory', err);
      this.inventoryItems = [];
      this.filteredItems = [];
      this.updatePagination();
      this.showToast('Failed to load inventory', 'error');
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
      this.availableSkus = await this.db.getSkusByCategory(this.newItem.category);
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

  onQuantityChange() {
    // Recalculate total required for current form item
    if (this.newItem.sku_code && this.newItem.qty) {
      // This will be calculated when materials are loaded
    }
  }

  async onFileSelected(event: Event) {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    this.uploadingMaster = true;
    this.uploadStatus = 'Uploading master data...';

    try {
      const result = await this.db.uploadMasterData(file);
      if (result.success) {
        this.uploadStatus = `Successfully imported ${result.count || 0} rows`;
        await this.loadCategories();
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
    return !!this.currentTable && // Must have a table selected
           !!this.newItem.category &&
           !!this.newItem.sku_code &&
           !!this.newItem.supplier?.trim() &&
           this.newItem.qty != null && this.newItem.qty > 0;
  }

  async addItem() {
    if (!this.canAddItem()) {
      this.showToast('Please fill all required fields and select a table', 'error');
      return;
    }

    this.addingItem = true;
    
    const userId = this.auth.getUserId();
    if (!userId) {
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
      qty: this.newItem.qty!,
      table_id: this.currentTable.id,
      user_id: userId
    };

    try {
      const res = await this.db.addInventoryItem(entry);
      if (res.success && res.id) {
        // Add to current table
        await this.addItemToTable(entry, res.id);
        
        const newItem: InventoryItem = { 
          id: res.id, 
          ...entry, 
          materialCount: 0,
          totalRequired: 0
        };
        
        this.inventoryItems.unshift(newItem);
        this.applyFilter();
        
        // Reset form
        this.newItem = { category: '', sku_code: '', sku_name: '', supplier: '', qty: null };
        this.availableSkus = [];
        
        this.showToast('Item added successfully', 'success');
      } else {
        this.showToast('Failed to save item', 'error');
      }
    } catch (err) {
      console.error('Add item error', err);
      this.showToast('Error adding item', 'error');
    } finally {
      this.addingItem = false;
    }
  }

  async addItemToTable(itemData: any, itemId: string) {
    if (!this.currentTable) return;
    
    try {
      const userId = this.auth.getUserId();
      if (!userId) {
        this.showToast('You must be logged in', 'error');
        return;
      }
      
      // Create requisition for the table with user_id for isolation
      const requisitionData = {
        table_id: this.currentTable.id,
        user_id: userId,
        sku_code: itemData.sku_code,
        sku_name: itemData.sku_name,
        category: itemData.category,
        supplier: itemData.supplier,
        qty_needed: itemData.qty,
        inventory_item_id: itemId
      };
      
      await this.db.createRequisition(requisitionData, []);
      
      // Update table item count
      const newCount = (this.currentTable.item_count || 0) + 1;
      await this.db.updateTableItemCount(this.currentTable.id, newCount, userId);
      this.currentTable.item_count = newCount;
      
      // Refresh tables list
      await this.loadUserTables();
    } catch (err) {
      console.error('Failed to add item to table', err);
      this.showToast('Failed to add item to table', 'error');
    }
  }

  async toggleRow(item: InventoryItem) {
    if (!item.id) return;
    
    this.expandedRows[item.id] = !this.expandedRows[item.id];

    if (this.expandedRows[item.id] && !item.materials) {
      this.loadingMaterials[item.id] = true;
      try {
        const materials = await this.db.getMaterialsForSku(item.sku_code);
        item.materials = materials || [];
        item.materialCount = item.materials?.length || 0;
        item.totalRequired = this.calculateTotalRequired(item);
      } catch (err) {
        console.error('Failed to load materials', err);
        item.materials = [];
        item.materialCount = 0;
        item.totalRequired = 0;
        this.showToast('Failed to load materials', 'error');
      } finally {
        this.loadingMaterials[item.id] = false;
      }
    }
  }

  applyFilter() {
    let list = [...this.inventoryItems];

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
    return this.filteredItems.reduce((sum, i) => sum + (i.qty || 0), 0);
  }

  // Calculation Methods
  calculateTotalRequired(item: InventoryItem): number {
    if (!item.materials || item.materials.length === 0) return 0;
    
    return item.materials.reduce((total, mat) => {
      const qtyPerBatch = mat.quantity_per_batch || 0;
      return total + (qtyPerBatch * (item.qty || 0));
    }, 0);
  }

  calculateMaterialTotal(itemQty: number | null, qtyPerBatch: number | null): number {
    const qty = itemQty || 0;
    const batchQty = qtyPerBatch || 0;
    return batchQty * qty;
  }

  // Table Management Methods
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

  async selectTable(table: UserTable) {
    this.currentTable = table;
    this.showTableDropdown = false;
    
    // Clear search/filters when switching tables
    this.searchQuery = '';
    this.filterCategory = '';
    
    // Load items for this table (only current user's data)
    await this.loadInventory();
    this.showToast(`Switched to table: ${table.name}`, 'info');
  }

  async createNewTable() {
    const tableName = prompt('Enter table name:');
    if (!tableName?.trim()) return;

    try {
      const userId = this.auth.getUserId();
      if (!userId) {
        this.showToast('You must be logged in to create tables', 'error');
        return;
      }

      const tableData = {
        user_id: userId,
        name: tableName.trim(),
        item_count: 0,
        created_at: new Date().toISOString(),
        updated_at: new Date().toISOString()
      };

      const result = await this.db.createUserTable(tableData);
      if (result.success && result.tableId) {
        await this.loadUserTables();
        this.showToast(`Table "${tableName}" created successfully`, 'success');
      } else {
        this.showToast('Failed to create table', 'error');
      }
    } catch (err) {
      console.error('Failed to create table', err);
      this.showToast('Error creating table', 'error');
    }
  }

  async renameTable(table: UserTable) {
    const newName = prompt('Enter new table name:', table.name);
    if (!newName?.trim() || newName === table.name) return;

    try {
      const userId = this.auth.getUserId();
      if (!userId) {
        this.showToast('You must be logged in', 'error');
        return;
      }

      const success = await this.db.updateTableName(table.id, newName.trim(), userId);
      if (success) {
        table.name = newName.trim();
        await this.loadUserTables();
        this.showToast('Table renamed successfully', 'success');
      } else {
        this.showToast('Failed to rename table', 'error');
      }
    } catch (err) {
      console.error('Failed to rename table', err);
      this.showToast('Error renaming table', 'error');
    }
  }

  async deleteTable(table: UserTable) {
    if (!confirm(`Are you sure you want to delete table "${table.name}"? This will also delete all items in this table. This action cannot be undone.`)) {
      return;
    }

    try {
      const userId = this.auth.getUserId();
      if (!userId) {
        this.showToast('You must be logged in', 'error');
        return;
      }

      const success = await this.db.deleteTable(table.id, userId);
      if (success) {
        // Remove from local array
        this.userTables = this.userTables.filter(t => t.id !== table.id);
        
        // If we're deleting the current table, select another or set to null
        if (this.currentTable?.id === table.id) {
          if (this.userTables.length > 0) {
            await this.selectTable(this.userTables[0]);
          } else {
            this.currentTable = null;
            this.inventoryItems = [];
            this.filteredItems = [];
            this.updatePagination();
          }
        }
        
        this.showToast(`Table "${table.name}" deleted successfully`, 'success');
      } else {
        this.showToast('Failed to delete table', 'error');
      }
    } catch (err) {
      console.error('Failed to delete table', err);
      this.showToast('Error deleting table', 'error');
    }
  }

  // Export Methods
  async exportData() {
    if (!this.currentTable) {
      this.showToast('No table selected', 'info');
      return;
    }

    if (this.filteredItems.length === 0) {
      this.showToast('No data to export', 'info');
      return;
    }

    try {
      const fileName = `${this.currentTable.name.replace(/\s+/g, '_')}_inventory_${new Date().toISOString().split('T')[0]}.csv`;
      
      // Create CSV content
      const headers = ['SKU Code', 'SKU Name', 'Category', 'Quantity', 'Supplier', 'Total Required', 'Materials Count'];
      const rows = this.filteredItems.map(item => [
        item.sku_code,
        item.sku_name || item.sku_code,
        item.category,
        item.qty.toString(),
        item.supplier || '',
        this.calculateTotalRequired(item).toString(),
        (item.materialCount || 0).toString()
      ]);

      const csvContent = [
        headers.join(','),
        ...rows.map(row => row.map(cell => `"${cell.replace(/"/g, '""')}"`).join(','))
      ].join('\n');

      // Create blob and download
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      
      link.setAttribute('href', url);
      link.setAttribute('download', fileName);
      link.style.visibility = 'hidden';
      
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      
      this.showToast(`Exported ${this.filteredItems.length} items from "${this.currentTable.name}" successfully`, 'success');
    } catch (err) {
      console.error('Export failed', err);
      this.showToast('Error exporting data', 'error');
    }
  }

  // Snackbar Methods
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

  // Pagination methods
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

  /**
   * Generate page numbers for pagination display
   * Shows: 1, 2, ..., current-1, current, current+1, ..., last-1, last
   * With ellipsis (...) for gaps
   */
  getPageNumbers(): number[] {
    const pages: number[] = [];
    const total = this.totalPages;
    const current = this.currentPage;
    const delta = 2; // Number of pages to show on each side of current page
    
    if (total <= 7) {
      // Show all pages if total is 7 or less
      for (let i = 1; i <= total; i++) {
        pages.push(i);
      }
    } else {
      // Always show first page
      pages.push(1);
      
      // Calculate start and end of page range around current page
      let start = Math.max(2, current - delta);
      let end = Math.min(total - 1, current + delta);
      
      // Add ellipsis before start if needed
      if (start > 2) {
        pages.push(-1); // -1 represents ellipsis
      }
      
      // Add page numbers in range
      for (let i = start; i <= end; i++) {
        pages.push(i);
      }
      
      // Add ellipsis after end if needed
      if (end < total - 1) {
        pages.push(-1); // -1 represents ellipsis
      }
      
      // Always show last page
      pages.push(total);
    }
    
    return pages;
  }

  /**
   * Navigate to specific page
   */
  goToPage(page: number) {
    if (page >= 1 && page <= this.totalPages && page !== this.currentPage) {
      this.currentPage = page;
      this.updatePagination();
    }
  }

  // Helper methods
  getTypeClass(type: string): string {
    if (!type) return '';
    const typeMap: { [key: string]: string } = {
      'conductor': 'type-conductor',
      'housing': 'type-housing',
      'electronics': 'type-electronics',
      'textile': 'type-textile',
      'sewing': 'type-sewing',
      'base-material': 'type-base',
      'printing': 'type-printing',
      'chemical': 'type-chemical',
      'binder': 'type-binder'
    };
    return typeMap[type.toLowerCase()] || '';
  }

  async deleteItem(item: InventoryItem, event: Event) {
    event.stopPropagation();
    
    if (!this.currentTable) {
      this.showToast('No table selected', 'error');
      return;
    }

    if (confirm('Are you sure you want to delete this item?')) {
      try {
        const userId = this.auth.getUserId();
        if (!userId) {
          this.showToast('You must be logged in', 'error');
          return;
        }

        const success = await this.db.deleteInventoryItem(item.id, userId, this.currentTable.id);
        
        if (success) {
          // Remove from local arrays
          this.inventoryItems = this.inventoryItems.filter(i => i.id !== item.id);
          this.applyFilter();
          
          // Update table item count
          const newCount = Math.max(0, (this.currentTable.item_count || 0) - 1);
          await this.db.updateTableItemCount(this.currentTable.id, newCount, userId);
          this.currentTable.item_count = newCount;
          
          // Refresh tables list
          await this.loadUserTables();
          
          this.showToast('Item deleted successfully', 'success');
        } else {
          this.showToast('Failed to delete item', 'error');
        }
      } catch (err) {
        console.error('Failed to delete item', err);
        this.showToast('Error deleting item', 'error');
      }
    }
  }
}