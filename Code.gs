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
 * Run this function ONCE after setting up your Google Apps Script project
 */
function initializePOSSystem() {
  try {
    console.log('🚀 Starting POS System initialization...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create Products sheet
    console.log('📦 Creating Products sheet...');
    createProductsSheet(ss);
    
    // Create Sales sheet
    console.log('💰 Creating Sales sheet...');
    createSalesSheet(ss);
    
    // Create Employees sheet
    console.log('👥 Creating Employees sheet...');
    createEmployeesSheet(ss);
    
    // Create Settings sheet
    console.log('⚙️ Creating Settings sheet...');
    createSettingsSheet(ss);
    
    // Create Logs sheet
    console.log('📋 Creating Logs sheet...');
    createLogsSheet(ss);
    
    // Create Sessions sheet
    console.log('⏰ Creating Sessions sheet...');
    createSessionsSheet(ss);
    
    // Initialize default data
    console.log('🔧 Adding default data...');
    initializeDefaultData();
    
    console.log('✅ POS system initialized successfully!');
    logAction('System Initialization', 'System', 'POS system initialized successfully');
    
    return '✅ SUCCESS: POS System initialized! Your database is ready. You can now deploy the web app.';
    
  } catch (error) {
    console.error('❌ Error initializing POS System:', error);
    return '❌ ERROR: Failed to initialize POS System: ' + error.toString();
  }
}

/**
 * Test function to verify system setup
 * Run this after initialization to verify everything works
 */
function testSystemSetup() {
  try {
    console.log('🔍 Testing system setup...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const sheetNames = sheets.map(sheet => sheet.getName());
    
    console.log('📊 Found sheets:', sheetNames);
    
    const requiredSheets = ['Products', 'Sales', 'Employees', 'Settings', 'Logs'];
    const missingSheets = requiredSheets.filter(name => !sheetNames.includes(name));
    
    if (missingSheets.length > 0) {
      return '❌ Missing sheets: ' + missingSheets.join(', ') + '. Please run initializePOSSystem() first.';
    }
    
    // Test user count
    const employeeSheet = ss.getSheetByName('Employees');
    const employeeCount = employeeSheet.getLastRow() - 1; // Subtract header row
    console.log('👥 Employee count:', employeeCount);
    
    // Test product count
    const productSheet = ss.getSheetByName('Products');
    const productCount = productSheet.getLastRow() - 1; // Subtract header row
    console.log('📦 Product count:', productCount);
    
    // Test authentication
    console.log('🔐 Testing authentication...');
    const authResult = authenticateUser('admin', '1234', 'pin');
    if (authResult.success) {
      console.log('✅ Authentication successful!');
      return `✅ SUCCESS: System setup complete!\n\n📊 Database Summary:\n- ${employeeCount} employees created\n- ${productCount} products added\n- All sheets configured\n- Admin authentication working\n\n🚀 Ready to deploy web app!`;
    } else {
      return '❌ Authentication failed: ' + authResult.message + '. Check employee data.';
    }
    
  } catch (error) {
    console.error('❌ Error testing system:', error);
    return '❌ ERROR: System test failed: ' + error.toString();
  }
}

/**
 * Reset the entire system (WARNING: This will delete all data!)
 */
function resetPOSSystem() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset POS System',
    'This will DELETE ALL DATA and reinitialize the system. Are you sure?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = ss.getSheets();
      
      // Delete all sheets except the first one
      for (let i = sheets.length - 1; i > 0; i--) {
        ss.deleteSheet(sheets[i]);
      }
      
      // Clear and rename the first sheet
      const firstSheet = sheets[0];
      firstSheet.clear();
      firstSheet.setName('Sheet1');
      
      // Reinitialize
      initializePOSSystem();
      
      return '✅ System reset and reinitialized successfully!';
    } catch (error) {
      return '❌ Error resetting system: ' + error.toString();
    }
  } else {
    return 'Reset cancelled.';
  }
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
  
  const headers = ['ID', 'Name', 'Role', 'Username', 'PIN Hash', 'Active', 'Created', 'Last Login', 'Password Hash'];
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
 * Create Logs sheet with comprehensive structure
 */
function createLogsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.LOGS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.LOGS);
  }
  
  sheet.clear();
  
  const headers = [
    'Timestamp', 'Action', 'User', 'Session ID', 'Details', 
    'Amount', 'Payment Method', 'Items Count', 'IP Address', 'Device Info'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Create Sessions sheet for tracking user sessions
 */
function createSessionsSheet(ss) {
  let sheet = ss.getSheetByName('Sessions');
  if (!sheet) {
    sheet = ss.insertSheet('Sessions');
  }
  
  sheet.clear();
  
  const headers = [
    'Session ID', 'User', 'Role', 'Opening Time', 'Closing Time', 
    'Duration (minutes)', 'Transactions Count', 'Total Sales', 'Status', 'IP Address'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1976d2');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Initialize default data for the system
 */
function initializeDefaultData() {
  // Add default admin user with both PIN and password
  addEmployee('Admin User', CONFIG.ROLES.ADMIN, 'admin', '1234', 'admin123');
  
  // Add sample employees
  addEmployee('John Manager', CONFIG.ROLES.MANAGER, 'john', '5678', 'john123');
  addEmployee('Sarah Cashier', CONFIG.ROLES.CASHIER, 'sarah', '9999', 'sarah123');
  
  // Add default settings
  updateSetting('TAX_RATE', '0.0875', 'Default tax rate (8.75%)');
  updateSetting('CURRENCY', 'USD', 'Currency symbol');
  updateSetting('LOW_STOCK_THRESHOLD', '10', 'Alert when stock falls below this number');
  updateSetting('RECEIPT_FOOTER', 'Thank you for your business!', 'Footer text for receipts');
  updateSetting('COMPANY_NAME', 'RetailPro', 'Company name for branding');
  updateSetting('COMPANY_TAGLINE', 'Modern Point of Sale System', 'Company tagline');
  
  // Add sample products
  addProduct('Coffee', 3.50, 1.20, 50, 'Beverages', '1234567890123');
  addProduct('Sandwich', 8.99, 4.50, 25, 'Food', '2345678901234');
  addProduct('Chips', 1.99, 0.75, 100, 'Snacks', '3456789012345');
  addProduct('Soda', 2.25, 0.90, 75, 'Beverages', '4567890123456');
  addProduct('Candy Bar', 1.50, 0.50, 200, 'Snacks', '5678901234567');
}

/**
 * Authentication Functions
 */

/**
 * Authenticate user login with PIN or Password
 */
function authenticateUser(username, credential, authType = 'pin') {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const [id, name, role, user, pinHash, active, created, lastLogin, passwordHash] = data[i];
      
      if (user === username && active === true) {
        let isValidCredential = false;
        
        if (authType === 'pin') {
          // PIN authentication
          const hashedPin = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, credential).toString();
          isValidCredential = (pinHash === hashedPin);
        } else if (authType === 'password') {
          // Password authentication
          const hashedPassword = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA256, credential).toString();
          isValidCredential = (passwordHash === hashedPassword);
        }
        
        if (isValidCredential) {
          // Update last login
          sheet.getRange(i + 1, 8).setValue(new Date());
          
          // Start user session
          const sessionId = startSession(username, role);
          
          logAction('User Login', username, `User ${name} logged in successfully using ${authType.toUpperCase()}`, sessionId);
          
          return {
            success: true,
            user: {
              id: id,
              name: name,
              role: role,
              username: user,
              sessionId: sessionId
            }
          };
        }
      }
    }
    
    logAction('Failed Login', username, `Invalid username or ${authType.toUpperCase()}`);
    return { success: false, message: `Invalid username or ${authType.toUpperCase()}` };
    
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
    
    logAction('Sale Completed', employeeId, `Sale completed: ${saleData.items.length} items, Total: $${totalSale.toFixed(2)}`, sessionId, totalSale, saleData.paymentMethod, saleData.items.length);
    
    return {
      success: true,
      sessionId: sessionId,
      total: totalSale,
      saleRecords: saleRecords.length,
      receipt: generateReceipt(sessionId, saleData, totalSale)
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
 * Add a new employee with PIN and optional password
 */
function addEmployee(name, role, username, pin, password = '') {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    const lastRow = sheet.getLastRow();
    const id = generateEmployeeId();
    const pinHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, pin).toString();
    const passwordHash = password ? Utilities.computeDigest(Utilities.DigestAlgorithm.SHA256, password).toString() : '';
    const now = new Date();
    
    const newEmployee = [id, name, role, username, pinHash, true, now, null, passwordHash];
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
function logAction(action, user, details, sessionId = null, amount = null, paymentMethod = null, itemsCount = null) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOGS);
    const lastRow = sheet.getLastRow();
    
    // Get client info (limited in Apps Script environment)
    const ipAddress = Session.getActiveUser().getEmail() || 'Unknown';
    const deviceInfo = 'Web Browser';
    
    const logEntry = [
      new Date(),
      action,
      user || 'System',
      sessionId || '',
      details || '',
      amount || '',
      paymentMethod || '',
      itemsCount || '',
      ipAddress,
      deviceInfo
    ];
    
    sheet.getRange(lastRow + 1, 1, 1, logEntry.length).setValues([logEntry]);
    
  } catch (error) {
    console.error('Log action error:', error);
  }
}

/**
 * Start user session tracking
 */
function startSession(user, role) {
  try {
    const sessionId = generateSessionId();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sessions');
    const now = new Date();
    const ipAddress = Session.getActiveUser().getEmail() || 'Unknown';
    
    const sessionEntry = [
      sessionId,
      user,
      role,
      now,
      '', // Closing time (empty for now)
      '', // Duration (calculated on close)
      0,  // Transactions count
      0,  // Total sales
      'ACTIVE',
      ipAddress
    ];
    
    sheet.appendRow(sessionEntry);
    
    // Log session start
    logAction('Session Start', user, `User ${user} started session`, sessionId);
    
    return sessionId;
    
  } catch (error) {
    console.error('Start session error:', error);
    return null;
  }
}

/**
 * End user session tracking
 */
function endSession(sessionId, user) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sessions');
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    
    // Find the session row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === sessionId && data[i][8] === 'ACTIVE') {
        const openingTime = new Date(data[i][3]);
        const duration = Math.round((now - openingTime) / (1000 * 60)); // Duration in minutes
        
        // Get session statistics
        const sessionStats = getSessionStatistics(sessionId);
        
        // Update the row
        sheet.getRange(i + 1, 5).setValue(now); // Closing time
        sheet.getRange(i + 1, 6).setValue(duration); // Duration
        sheet.getRange(i + 1, 7).setValue(sessionStats.transactionCount); // Transactions
        sheet.getRange(i + 1, 8).setValue(sessionStats.totalSales); // Total sales
        sheet.getRange(i + 1, 9).setValue('CLOSED'); // Status
        
        // Log session end
        logAction('Session End', user, `Session duration: ${duration} minutes, Transactions: ${sessionStats.transactionCount}, Sales: $${sessionStats.totalSales.toFixed(2)}`, sessionId);
        
        return {
          duration: duration,
          transactions: sessionStats.transactionCount,
          totalSales: sessionStats.totalSales
        };
      }
    }
    
  } catch (error) {
    console.error('End session error:', error);
    return null;
  }
}

/**
 * Get session statistics
 */
function getSessionStatistics(sessionId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOGS);
    const data = sheet.getDataRange().getValues();
    
    let transactionCount = 0;
    let totalSales = 0;
    
    // Count transactions and sum sales for this session
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === sessionId && data[i][1] === 'Sale Completed') {
        transactionCount++;
        const amount = parseFloat(data[i][5]) || 0;
        totalSales += amount;
      }
    }
    
    return {
      transactionCount: transactionCount,
      totalSales: totalSales
    };
    
  } catch (error) {
    console.error('Get session statistics error:', error);
    return { transactionCount: 0, totalSales: 0 };
  }
}

/**
 * Generate unique session ID
 */
function generateSessionId() {
  const timestamp = new Date().getTime();
  const random = Math.floor(Math.random() * 1000);
  return `SES${timestamp}${random}`;
}

/**
 * Get current user (placeholder - implement based on your auth system)
 */
function getCurrentUser() {
  // This would typically get the current authenticated user
  return Session.getActiveUser().getEmail() || 'Unknown';
}

/**
 * Get comprehensive sales reports
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
      totalCost: 0,
      totalProfit: 0,
      paymentMethods: {},
      topProducts: {},
      employeePerformance: {},
      hourlyBreakdown: {},
      categoryBreakdown: {}
    };
    
    for (let i = 1; i < data.length; i++) {
      const [saleId, date, productId, productName, quantity, unitPrice, total, tax, discount, employee, paymentMethod] = data[i];
      
      const saleDate = new Date(date);
      if ((!startDate || saleDate >= startDate) && (!endDate || saleDate <= endDate)) {
        // Get product cost for profit calculation
        const productCost = getProductCost(productId) || 0;
        const itemCost = productCost * quantity;
        const itemProfit = (unitPrice - productCost) * quantity;
        
        sales.push({
          saleId, date, productId, productName, quantity, unitPrice, total, tax, discount, employee, paymentMethod,
          cost: itemCost, profit: itemProfit
        });
        
        summary.totalSales += total;
        summary.totalItems += quantity;
        summary.totalTax += tax;
        summary.totalDiscount += discount;
        summary.totalCost += itemCost;
        summary.totalProfit += itemProfit;
        
        // Payment methods breakdown
        summary.paymentMethods[paymentMethod] = (summary.paymentMethods[paymentMethod] || 0) + total;
        
        // Top products by quantity
        summary.topProducts[productName] = (summary.topProducts[productName] || 0) + quantity;
        
        // Employee performance
        if (!summary.employeePerformance[employee]) {
          summary.employeePerformance[employee] = { sales: 0, items: 0, transactions: 0 };
        }
        summary.employeePerformance[employee].sales += total;
        summary.employeePerformance[employee].items += quantity;
        summary.employeePerformance[employee].transactions += 1;
        
        // Hourly breakdown
        const hour = saleDate.getHours();
        const hourKey = `${hour}:00-${hour + 1}:00`;
        summary.hourlyBreakdown[hourKey] = (summary.hourlyBreakdown[hourKey] || 0) + total;
        
        // Category breakdown
        const category = getProductCategory(productId) || 'Unknown';
        summary.categoryBreakdown[category] = (summary.categoryBreakdown[category] || 0) + total;
      }
    }
    
    return { success: true, sales: sales, summary: summary };
    
  } catch (error) {
    console.error('Get sales reports error:', error);
    return { success: false, message: 'Failed to get sales reports' };
  }
}

/**
 * Get Daily RVC (Revenue, Volume, Cost) Report
 */
function getDailyRVCReport(targetDate) {
  try {
    const startDate = new Date(targetDate);
    startDate.setHours(0, 0, 0, 0);
    const endDate = new Date(targetDate);
    endDate.setHours(23, 59, 59, 999);
    
    const salesResult = getSalesReports(startDate, endDate);
    if (!salesResult.success) return salesResult;
    
    const { sales, summary } = salesResult;
    
    // Calculate additional metrics
    const totalRevenue = summary.totalSales;
    const totalVolume = summary.totalItems;
    const totalCost = summary.totalCost;
    const totalProfit = summary.totalProfit;
    const profitMargin = totalRevenue > 0 ? (totalProfit / totalRevenue) * 100 : 0;
    const avgTransactionValue = sales.length > 0 ? totalRevenue / getUniqueTransactions(sales) : 0;
    
    // Hourly breakdown
    const hourlyData = Array.from({ length: 24 }, (_, i) => {
      const hourKey = `${i}:00-${i + 1}:00`;
      return {
        hour: hourKey,
        revenue: summary.hourlyBreakdown[hourKey] || 0,
        transactions: getHourlyTransactions(sales, i)
      };
    });
    
    // Top performing products
    const topProductsByRevenue = getTopProductsByRevenue(sales);
    const topProductsByVolume = Object.entries(summary.topProducts)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 10)
      .map(([name, quantity]) => ({ name, quantity }));
    
    return {
      success: true,
      report: {
        date: targetDate,
        summary: {
          totalRevenue,
          totalVolume,
          totalCost,
          totalProfit,
          profitMargin,
          avgTransactionValue,
          totalTransactions: getUniqueTransactions(sales),
          totalTax: summary.totalTax,
          totalDiscount: summary.totalDiscount
        },
        hourlyBreakdown: hourlyData,
        topProductsByRevenue,
        topProductsByVolume,
        paymentMethods: summary.paymentMethods,
        employeePerformance: summary.employeePerformance,
        categoryBreakdown: summary.categoryBreakdown
      }
    };
    
  } catch (error) {
    console.error('Daily RVC report error:', error);
    return { success: false, message: 'Failed to generate daily RVC report' };
  }
}

/**
 * Get Weekly Sales Summary Report
 */
function getWeeklySalesReport(weekStartDate) {
  try {
    const startDate = new Date(weekStartDate);
    const endDate = new Date(weekStartDate);
    endDate.setDate(endDate.getDate() + 6);
    endDate.setHours(23, 59, 59, 999);
    
    const salesResult = getSalesReports(startDate, endDate);
    if (!salesResult.success) return salesResult;
    
    // Group by days
    const dailyBreakdown = {};
    const { sales } = salesResult;
    
    sales.forEach(sale => {
      const dateKey = new Date(sale.date).toDateString();
      if (!dailyBreakdown[dateKey]) {
        dailyBreakdown[dateKey] = { revenue: 0, transactions: 0, items: 0 };
      }
      dailyBreakdown[dateKey].revenue += sale.total;
      dailyBreakdown[dateKey].items += sale.quantity;
    });
    
    // Count unique transactions per day
    const transactionsByDay = {};
    sales.forEach(sale => {
      const dateKey = new Date(sale.date).toDateString();
      if (!transactionsByDay[dateKey]) {
        transactionsByDay[dateKey] = new Set();
      }
      transactionsByDay[dateKey].add(sale.saleId.split('_')[0]); // Get session ID
    });
    
    Object.keys(transactionsByDay).forEach(dateKey => {
      if (dailyBreakdown[dateKey]) {
        dailyBreakdown[dateKey].transactions = transactionsByDay[dateKey].size;
      }
    });
    
    return {
      success: true,
      report: {
        weekStart: weekStartDate,
        weekEnd: endDate,
        dailyBreakdown,
        summary: salesResult.summary
      }
    };
    
  } catch (error) {
    console.error('Weekly sales report error:', error);
    return { success: false, message: 'Failed to generate weekly sales report' };
  }
}

/**
 * Get Monthly Performance Report
 */
function getMonthlyPerformanceReport(year, month) {
  try {
    const startDate = new Date(year, month - 1, 1);
    const endDate = new Date(year, month, 0, 23, 59, 59, 999);
    
    const salesResult = getSalesReports(startDate, endDate);
    if (!salesResult.success) return salesResult;
    
    const { sales, summary } = salesResult;
    
    // Weekly breakdown
    const weeklyData = [];
    let currentWeekStart = new Date(startDate);
    
    while (currentWeekStart <= endDate) {
      const weekEnd = new Date(currentWeekStart);
      weekEnd.setDate(weekEnd.getDate() + 6);
      if (weekEnd > endDate) weekEnd.setTime(endDate.getTime());
      
      const weekSales = sales.filter(sale => {
        const saleDate = new Date(sale.date);
        return saleDate >= currentWeekStart && saleDate <= weekEnd;
      });
      
      const weekRevenue = weekSales.reduce((sum, sale) => sum + sale.total, 0);
      const weekItems = weekSales.reduce((sum, sale) => sum + sale.quantity, 0);
      
      weeklyData.push({
        weekStart: new Date(currentWeekStart),
        weekEnd: new Date(weekEnd),
        revenue: weekRevenue,
        items: weekItems,
        transactions: getUniqueTransactions(weekSales)
      });
      
      currentWeekStart.setDate(currentWeekStart.getDate() + 7);
    }
    
    return {
      success: true,
      report: {
        year,
        month,
        monthName: new Date(year, month - 1).toLocaleString('default', { month: 'long' }),
        weeklyBreakdown: weeklyData,
        summary: summary,
        trends: calculateTrends(weeklyData)
      }
    };
    
  } catch (error) {
    console.error('Monthly performance report error:', error);
    return { success: false, message: 'Failed to generate monthly performance report' };
  }
}

/**
 * Get Product Performance Report
 */
function getProductPerformanceReport(startDate, endDate) {
  try {
    const salesResult = getSalesReports(startDate, endDate);
    if (!salesResult.success) return salesResult;
    
    const { sales } = salesResult;
    const productStats = {};
    
    sales.forEach(sale => {
      if (!productStats[sale.productId]) {
        productStats[sale.productId] = {
          id: sale.productId,
          name: sale.productName,
          totalRevenue: 0,
          totalQuantity: 0,
          totalCost: 0,
          totalProfit: 0,
          transactions: new Set()
        };
      }
      
      const stats = productStats[sale.productId];
      stats.totalRevenue += sale.total;
      stats.totalQuantity += sale.quantity;
      stats.totalCost += sale.cost || 0;
      stats.totalProfit += sale.profit || 0;
      stats.transactions.add(sale.saleId.split('_')[0]);
    });
    
    // Convert to array and add calculated metrics
    const productPerformance = Object.values(productStats).map(product => ({
      ...product,
      avgSellingPrice: product.totalQuantity > 0 ? product.totalRevenue / product.totalQuantity : 0,
      profitMargin: product.totalRevenue > 0 ? (product.totalProfit / product.totalRevenue) * 100 : 0,
      transactions: product.transactions.size
    }));
    
    // Sort by revenue
    productPerformance.sort((a, b) => b.totalRevenue - a.totalRevenue);
    
    return {
      success: true,
      report: {
        startDate,
        endDate,
        products: productPerformance,
        topByRevenue: productPerformance.slice(0, 10),
        topByQuantity: [...productPerformance].sort((a, b) => b.totalQuantity - a.totalQuantity).slice(0, 10),
        topByProfit: [...productPerformance].sort((a, b) => b.totalProfit - a.totalProfit).slice(0, 10)
      }
    };
    
  } catch (error) {
    console.error('Product performance report error:', error);
    return { success: false, message: 'Failed to generate product performance report' };
  }
}

/**
 * Get Employee Performance Report
 */
function getEmployeePerformanceReport(startDate, endDate) {
  try {
    const salesResult = getSalesReports(startDate, endDate);
    if (!salesResult.success) return salesResult;
    
    const { summary } = salesResult;
    const employeeStats = [];
    
    Object.entries(summary.employeePerformance).forEach(([employee, stats]) => {
      employeeStats.push({
        employee,
        totalSales: stats.sales,
        totalItems: stats.items,
        totalTransactions: stats.transactions,
        avgTransactionValue: stats.transactions > 0 ? stats.sales / stats.transactions : 0,
        avgItemsPerTransaction: stats.transactions > 0 ? stats.items / stats.transactions : 0
      });
    });
    
    // Sort by total sales
    employeeStats.sort((a, b) => b.totalSales - a.totalSales);
    
    return {
      success: true,
      report: {
        startDate,
        endDate,
        employees: employeeStats,
        topPerformers: employeeStats.slice(0, 5)
      }
    };
    
  } catch (error) {
    console.error('Employee performance report error:', error);
    return { success: false, message: 'Failed to generate employee performance report' };
  }
}

/**
 * Helper function to get product cost
 */
function getProductCost(productId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PRODUCTS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productId) {
        return data[i][3]; // Cost column
      }
    }
    return 0;
  } catch (error) {
    return 0;
  }
}

/**
 * Helper function to get product category
 */
function getProductCategory(productId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PRODUCTS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productId) {
        return data[i][5]; // Category column
      }
    }
    return 'Unknown';
  } catch (error) {
    return 'Unknown';
  }
}

/**
 * Helper function to get unique transactions count
 */
function getUniqueTransactions(sales) {
  const sessions = new Set();
  sales.forEach(sale => {
    const sessionId = sale.saleId.split('_')[0];
    sessions.add(sessionId);
  });
  return sessions.size;
}

/**
 * Helper function to get hourly transactions
 */
function getHourlyTransactions(sales, hour) {
  const hourSales = sales.filter(sale => {
    const saleHour = new Date(sale.date).getHours();
    return saleHour === hour;
  });
  return getUniqueTransactions(hourSales);
}

/**
 * Helper function to get top products by revenue
 */
function getTopProductsByRevenue(sales) {
  const productRevenue = {};
  
  sales.forEach(sale => {
    if (!productRevenue[sale.productName]) {
      productRevenue[sale.productName] = 0;
    }
    productRevenue[sale.productName] += sale.total;
  });
  
  return Object.entries(productRevenue)
    .sort(([,a], [,b]) => b - a)
    .slice(0, 10)
    .map(([name, revenue]) => ({ name, revenue }));
}

/**
 * Helper function to calculate trends
 */
function calculateTrends(weeklyData) {
  if (weeklyData.length < 2) return { revenue: 0, items: 0, transactions: 0 };
  
  const current = weeklyData[weeklyData.length - 1];
  const previous = weeklyData[weeklyData.length - 2];
  
  return {
    revenue: previous.revenue > 0 ? ((current.revenue - previous.revenue) / previous.revenue) * 100 : 0,
    items: previous.items > 0 ? ((current.items - previous.items) / previous.items) * 100 : 0,
    transactions: previous.transactions > 0 ? ((current.transactions - previous.transactions) / previous.transactions) * 100 : 0
  };
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

/**
 * Generate professional receipt
 */
function generateReceipt(sessionId, saleData, total) {
  try {
    const companyName = getSetting('COMPANY_NAME') || 'RetailPro';
    const companyTagline = getSetting('COMPANY_TAGLINE') || 'Modern Point of Sale System';
    const currency = getSetting('CURRENCY') || 'USD';
    const taxRate = parseFloat(getSetting('TAX_RATE') || '0.0875');
    
    const now = new Date();
    const subtotal = total / (1 + taxRate);
    const tax = total - subtotal;
    const discount = saleData.discount || 0;
    
    const receipt = {
      // Header
      companyName: companyName,
      companyTagline: companyTagline,
      
      // Transaction Details
      receiptNumber: sessionId,
      date: now.toLocaleDateString(),
      time: now.toLocaleTimeString(),
      cashier: saleData.employee,
      
      // Items
      items: saleData.items.map(item => ({
        name: item.name,
        quantity: item.quantity,
        unitPrice: item.price,
        total: item.quantity * item.price
      })),
      
      // Totals
      subtotal: subtotal,
      tax: tax,
      taxRate: (taxRate * 100).toFixed(2),
      discount: discount,
      total: total,
      currency: currency,
      
      // Payment
      paymentMethod: saleData.paymentMethod,
      amountPaid: saleData.amountPaid || total,
      change: (saleData.amountPaid || total) - total,
      
      // Footer
      footerMessage: getSetting('RECEIPT_FOOTER') || 'Thank you for your business!',
      
      // Session Info
      sessionId: sessionId,
      transactionId: saleData.transactionId || sessionId
    };
    
    return receipt;
    
  } catch (error) {
    console.error('Generate receipt error:', error);
    return null;
  }
}

/**
 * Create daily backup
 */
function createDailyBackup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const backupName = `POS_Backup_${new Date().toISOString().split('T')[0]}`;
    
    // Create a copy of the spreadsheet
    const backup = ss.copy(backupName);
    
    // Move to a backup folder if exists
    try {
      const backupFolder = DriveApp.getFoldersByName('POS_Backups').next();
      const file = DriveApp.getFileById(backup.getId());
      file.moveTo(backupFolder);
    } catch (e) {
      // Create backup folder if it doesn't exist
      const backupFolder = DriveApp.createFolder('POS_Backups');
      const file = DriveApp.getFileById(backup.getId());
      file.moveTo(backupFolder);
    }
    
    logAction('System Backup', 'System', `Daily backup created: ${backupName}`);
    
    return {
      success: true,
      backupId: backup.getId(),
      backupName: backupName
    };
    
  } catch (error) {
    console.error('Create backup error:', error);
    logAction('System Error', 'System', `Backup failed: ${error.toString()}`);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Schedule daily backup (can be set up as a trigger)
 */
function scheduleDailyBackup() {
  // Delete existing daily triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'createDailyBackup') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new daily trigger at 2 AM
  ScriptApp.newTrigger('createDailyBackup')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
    
  logAction('System Setup', 'System', 'Daily backup scheduled at 2:00 AM');
}

/**
 * End user session (to be called on logout)
 */
function logoutUser(sessionId, username) {
  try {
    if (sessionId) {
      const sessionSummary = endSession(sessionId, username);
      logAction('User Logout', username, `User logged out. Session summary: ${JSON.stringify(sessionSummary)}`, sessionId);
      return sessionSummary;
    }
    return null;
  } catch (error) {
    console.error('Logout error:', error);
    return null;
  }
}