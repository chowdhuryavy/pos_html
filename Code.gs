/**
 * Modern POS System - Google Apps Script Backend
 * Handles all database operations, authentication, and business logic
 */

// Configuration constants
const CONFIG = {
  SHEETS: {
    PRODUCTS: 'Products',
    SALES: 'Sales', 
    EMPLOYEES: 'Employees',
    SETTINGS: 'Settings',
    LOGS: 'Logs'
  },
  ROLES: {
    ADMIN: 'Admin',
    MANAGER: 'Manager', 
    CASHIER: 'Cashier'
  }
};

/**
 * Initialize the POS system by creating necessary sheets
 */
function initializePOSSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Products sheet
  createProductsSheet(ss);
  
  // Create Sales sheet
  createSalesSheet(ss);
  
  // Create Employees sheet
  createEmployeesSheet(ss);
  
  // Create Settings sheet
  createSettingsSheet(ss);
  
  // Create Logs sheet
  createLogsSheet(ss);
  
  // Initialize default data
  initializeDefaultData();
  
  logAction('System Initialization', 'System', 'POS system initialized successfully');
}

/**
 * Create Products sheet with proper structure
 */
function createProductsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.PRODUCTS);
  }
  
  // Clear existing content
  sheet.clear();
  
  // Set headers
  const headers = ['ID', 'Name', 'Price', 'Cost', 'Stock', 'Category', 'Barcode', 'Active', 'Created', 'Updated'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  // Freeze header row
  sheet.setFrozenRows(1);
}

/**
 * Create Sales sheet with proper structure
 */
function createSalesSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SALES);
  }
  
  sheet.clear();
  
  const headers = ['Sale ID', 'Date', 'Product ID', 'Product Name', 'Quantity', 'Unit Price', 'Total', 'Tax', 'Discount', 'Employee', 'Payment Method', 'Session ID'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
}

/**
 * Create Employees sheet with proper structure
 */
function createEmployeesSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.EMPLOYEES);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.EMPLOYEES);
  }
  
  sheet.clear();
  
  const headers = ['ID', 'Name', 'Role', 'Username', 'PIN Hash', 'Active', 'Created', 'Last Login'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#ea4335');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
}

/**
 * Create Settings sheet with proper structure
 */
function createSettingsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SETTINGS);
  }
  
  sheet.clear();
  
  const headers = ['Setting', 'Value', 'Description'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#fbbc04');
  headerRange.setFontColor('black');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
}

/**
 * Create Logs sheet with proper structure
 */
function createLogsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.LOGS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.LOGS);
  }
  
  sheet.clear();
  
  const headers = ['Timestamp', 'Action', 'User', 'Details', 'IP'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#9aa0a6');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
}

/**
 * Initialize default data for the system
 */
function initializeDefaultData() {
  // Add default admin user
  addEmployee('Admin User', CONFIG.ROLES.ADMIN, 'admin', '1234');
  
  // Add default settings
  updateSetting('TAX_RATE', '0.0875', 'Default tax rate (8.75%)');
  updateSetting('CURRENCY', 'USD', 'Currency symbol');
  updateSetting('LOW_STOCK_THRESHOLD', '10', 'Alert when stock falls below this number');
  updateSetting('RECEIPT_FOOTER', 'Thank you for your business!', 'Footer text for receipts');
  
  // Add sample products
  addProduct('Coffee', 3.50, 1.20, 50, 'Beverages', '1234567890123');
  addProduct('Sandwich', 8.99, 4.50, 25, 'Food', '2345678901234');
  addProduct('Chips', 1.99, 0.75, 100, 'Snacks', '3456789012345');
}

/**
 * Authentication Functions
 */

/**
 * Authenticate user login
 */
function authenticateUser(username, pin) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const [id, name, role, user, pinHash, active] = data[i];
      
      if (user === username && active === true) {
        // For demo purposes, we'll use simple hash comparison
        const hashedPin = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, pin).toString();
        
        if (pinHash === hashedPin) {
          // Update last login
          sheet.getRange(i + 1, 8).setValue(new Date());
          
          logAction('User Login', username, `User ${name} logged in successfully`);
          
          return {
            success: true,
            user: {
              id: id,
              name: name,
              role: role,
              username: user
            }
          };
        }
      }
    }
    
    logAction('Failed Login', username, 'Invalid username or PIN');
    return { success: false, message: 'Invalid username or PIN' };
    
  } catch (error) {
    console.error('Authentication error:', error);
    return { success: false, message: 'Authentication failed' };
  }
}

/**
 * Product Management Functions
 */

/**
 * Add a new product
 */
function addProduct(name, price, cost, stock, category, barcode) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PRODUCTS);
    const lastRow = sheet.getLastRow();
    const id = generateProductId();
    const now = new Date();
    
    const newProduct = [id, name, price, cost, stock, category, barcode, true, now, now];
    sheet.getRange(lastRow + 1, 1, 1, newProduct.length).setValues([newProduct]);
    
    logAction('Product Added', getCurrentUser(), `Added product: ${name}`);
    return { success: true, id: id };
    
  } catch (error) {
    console.error('Add product error:', error);
    return { success: false, message: 'Failed to add product' };
  }
}

/**
 * Get all active products
 */
function getProducts() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PRODUCTS);
    const data = sheet.getDataRange().getValues();
    const products = [];
    
    for (let i = 1; i < data.length; i++) {
      const [id, name, price, cost, stock, category, barcode, active] = data[i];
      
      if (active) {
        products.push({
          id: id,
          name: name,
          price: price,
          cost: cost,
          stock: stock,
          category: category,
          barcode: barcode
        });
      }
    }
    
    return { success: true, products: products };
    
  } catch (error) {
    console.error('Get products error:', error);
    return { success: false, message: 'Failed to get products' };
  }
}

/**
 * Search products by name or barcode
 */
function searchProducts(query) {
  try {
    const result = getProducts();
    if (!result.success) return result;
    
    const filteredProducts = result.products.filter(product => 
      product.name.toLowerCase().includes(query.toLowerCase()) ||
      product.barcode.includes(query) ||
      product.category.toLowerCase().includes(query.toLowerCase())
    );
    
    return { success: true, products: filteredProducts };
    
  } catch (error) {
    console.error('Search products error:', error);
    return { success: false, message: 'Failed to search products' };
  }
}

/**
 * Sales Functions
 */

/**
 * Process a sale
 */
function processSale(saleData) {
  try {
    const { items, paymentMethod, employeeId, taxRate, discount } = saleData;
    const sessionId = generateSessionId();
    const saleDate = new Date();
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SALES);
    
    let totalSale = 0;
    const saleRecords = [];
    
    // Process each item
    for (const item of items) {
      const saleId = generateSaleId();
      const itemTotal = item.quantity * item.price;
      const itemTax = itemTotal * (taxRate || 0);
      const itemDiscount = itemTotal * (discount || 0);
      
      totalSale += itemTotal + itemTax - itemDiscount;
      
      // Update stock
      updateProductStock(item.id, -item.quantity);
      
      saleRecords.push([
        saleId,
        saleDate,
        item.id,
        item.name,
        item.quantity,
        item.price,
        itemTotal,
        itemTax,
        itemDiscount,
        employeeId,
        paymentMethod,
        sessionId
      ]);
    }
    
    // Add all sale records
    if (saleRecords.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, saleRecords.length, saleRecords[0].length).setValues(saleRecords);
    }
    
    logAction('Sale Processed', employeeId, `Sale processed for $${totalSale.toFixed(2)}`);
    
    return {
      success: true,
      sessionId: sessionId,
      total: totalSale,
      saleRecords: saleRecords.length
    };
    
  } catch (error) {
    console.error('Process sale error:', error);
    return { success: false, message: 'Failed to process sale' };
  }
}

/**
 * Employee Management Functions
 */

/**
 * Add a new employee
 */
function addEmployee(name, role, username, pin) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    const lastRow = sheet.getLastRow();
    const id = generateEmployeeId();
    const pinHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, pin).toString();
    const now = new Date();
    
    const newEmployee = [id, name, role, username, pinHash, true, now, null];
    sheet.getRange(lastRow + 1, 1, 1, newEmployee.length).setValues([newEmployee]);
    
    logAction('Employee Added', getCurrentUser(), `Added employee: ${name}`);
    return { success: true, id: id };
    
  } catch (error) {
    console.error('Add employee error:', error);
    return { success: false, message: 'Failed to add employee' };
  }
}

/**
 * Utility Functions
 */

/**
 * Generate unique product ID
 */
function generateProductId() {
  return 'P' + Date.now().toString().slice(-8);
}

/**
 * Generate unique sale ID
 */
function generateSaleId() {
  return 'S' + Date.now().toString() + Math.random().toString(36).substr(2, 5);
}

/**
 * Generate unique employee ID
 */
function generateEmployeeId() {
  return 'E' + Date.now().toString().slice(-8);
}

/**
 * Generate unique session ID
 */
function generateSessionId() {
  return 'SESSION_' + Date.now().toString() + '_' + Math.random().toString(36).substr(2, 5);
}

/**
 * Update product stock
 */
function updateProductStock(productId, quantity) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PRODUCTS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productId) {
        const currentStock = data[i][4];
        const newStock = currentStock + quantity;
        sheet.getRange(i + 1, 5).setValue(newStock);
        sheet.getRange(i + 1, 10).setValue(new Date()); // Update timestamp
        
        // Check for low stock
        const lowStockThreshold = getSetting('LOW_STOCK_THRESHOLD') || 10;
        if (newStock <= lowStockThreshold) {
          logAction('Low Stock Alert', 'System', `Product ${data[i][1]} is low in stock: ${newStock} remaining`);
        }
        
        return { success: true, newStock: newStock };
      }
    }
    
    return { success: false, message: 'Product not found' };
    
  } catch (error) {
    console.error('Update stock error:', error);
    return { success: false, message: 'Failed to update stock' };
  }
}

/**
 * Update system setting
 */
function updateSetting(key, value, description) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SETTINGS);
    const data = sheet.getDataRange().getValues();
    
    // Check if setting exists
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        return { success: true };
      }
    }
    
    // Add new setting
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, 3).setValues([[key, value, description]]);
    
    return { success: true };
    
  } catch (error) {
    console.error('Update setting error:', error);
    return { success: false, message: 'Failed to update setting' };
  }
}

/**
 * Get system setting
 */
function getSetting(key) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SETTINGS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        return data[i][1];
      }
    }
    
    return null;
    
  } catch (error) {
    console.error('Get setting error:', error);
    return null;
  }
}

/**
 * Log system actions
 */
function logAction(action, user, details) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOGS);
    const lastRow = sheet.getLastRow();
    
    const logEntry = [
      new Date(),
      action,
      user || 'System',
      details,
      Session.getActiveUser().getEmail() || 'Unknown'
    ];
    
    sheet.getRange(lastRow + 1, 1, 1, logEntry.length).setValues([logEntry]);
    
  } catch (error) {
    console.error('Log action error:', error);
  }
}

/**
 * Get current user (placeholder - implement based on your auth system)
 */
function getCurrentUser() {
  // This would typically get the current authenticated user
  return Session.getActiveUser().getEmail() || 'Unknown';
}

/**
 * Get sales reports
 */
function getSalesReports(startDate, endDate) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SALES);
    const data = sheet.getDataRange().getValues();
    
    const sales = [];
    const summary = {
      totalSales: 0,
      totalItems: 0,
      totalTax: 0,
      totalDiscount: 0,
      paymentMethods: {},
      topProducts: {}
    };
    
    for (let i = 1; i < data.length; i++) {
      const [saleId, date, productId, productName, quantity, unitPrice, total, tax, discount, employee, paymentMethod] = data[i];
      
      const saleDate = new Date(date);
      if ((!startDate || saleDate >= startDate) && (!endDate || saleDate <= endDate)) {
        sales.push({
          saleId, date, productId, productName, quantity, unitPrice, total, tax, discount, employee, paymentMethod
        });
        
        summary.totalSales += total;
        summary.totalItems += quantity;
        summary.totalTax += tax;
        summary.totalDiscount += discount;
        
        summary.paymentMethods[paymentMethod] = (summary.paymentMethods[paymentMethod] || 0) + total;
        summary.topProducts[productName] = (summary.topProducts[productName] || 0) + quantity;
      }
    }
    
    return { success: true, sales: sales, summary: summary };
    
  } catch (error) {
    console.error('Get sales reports error:', error);
    return { success: false, message: 'Failed to get sales reports' };
  }
}

/**
 * Web app entry points
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Include HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}