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
    
    // Create Enterprise sheets
    console.log('🏪 Creating Tables sheet...');
    createTablesSheet(ss);
    
    console.log('📋 Creating Menu Categories sheet...');
    createMenuCategoriesSheet(ss);
    
    console.log('👨‍🍳 Creating Kitchen Orders sheet...');
    createKitchenOrdersSheet(ss);
    
    console.log('📦 Creating Inventory Adjustments sheet...');
    createInventoryAdjustmentsSheet(ss);
    
    console.log('🎯 Creating Promotions sheet...');
    createPromotionsSheet(ss);
    
    console.log('⏰ Creating Time Clock sheet...');
    createTimeClockSheet(ss);
    
    // Create Admin/Permissions sheets
    console.log('🔐 Creating Permissions sheet...');
    createPermissionsSheet(ss);
    
    console.log('👥 Creating Roles sheet...');
    createRolesSheet(ss);
    
    // Create Advanced POS sheets
    console.log('👤 Creating Customers sheet...');
    createCustomersSheet(ss);
    
    console.log('💳 Creating Split Payments sheet...');
    createSplitPaymentsSheet(ss);
    
    console.log('🎁 Creating Loyalty Program sheet...');
    createLoyaltyProgramSheet(ss);
    
    console.log('🏷️ Creating Advanced Discounts sheet...');
    createAdvancedDiscountsSheet(ss);
    
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
  updateSetting('COMPANY_LOGO', 'fas fa-store', 'FontAwesome icon class for logo');
  updateSetting('COMPANY_ADDRESS', '123 Business St, City, State 12345', 'Company address for receipts');
  updateSetting('COMPANY_PHONE', '(555) 123-4567', 'Company phone number');
  updateSetting('COMPANY_EMAIL', 'info@retailpro.com', 'Company email address');
  updateSetting('COMPANY_WEBSITE', 'www.retailpro.com', 'Company website');
  updateSetting('BUSINESS_HOURS', 'Mon-Fri: 9AM-9PM, Sat-Sun: 10AM-8PM', 'Business hours for receipts');
  
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
 * Get all company branding settings
 */
function getCompanySettings() {
  try {
    return {
      name: getSetting('COMPANY_NAME') || 'RetailPro',
      tagline: getSetting('COMPANY_TAGLINE') || 'Modern Point of Sale System',
      logo: getSetting('COMPANY_LOGO') || 'fas fa-store',
      address: getSetting('COMPANY_ADDRESS') || '123 Business St, City, State 12345',
      phone: getSetting('COMPANY_PHONE') || '(555) 123-4567',
      email: getSetting('COMPANY_EMAIL') || 'info@retailpro.com',
      website: getSetting('COMPANY_WEBSITE') || 'www.retailpro.com',
      businessHours: getSetting('BUSINESS_HOURS') || 'Mon-Fri: 9AM-9PM, Sat-Sun: 10AM-8PM',
      taxRate: parseFloat(getSetting('TAX_RATE') || '0.0875'),
      currency: getSetting('CURRENCY') || 'USD',
      receiptFooter: getSetting('RECEIPT_FOOTER') || 'Thank you for your business!'
    };
  } catch (error) {
    console.error('Get company settings error:', error);
    return {
      name: 'RetailPro',
      tagline: 'Modern Point of Sale System',
      logo: 'fas fa-store',
      address: '123 Business St, City, State 12345',
      phone: '(555) 123-4567',
      email: 'info@retailpro.com',
      website: 'www.retailpro.com',
      businessHours: 'Mon-Fri: 9AM-9PM, Sat-Sun: 10AM-8PM',
      taxRate: 0.0875,
      currency: 'USD',
      receiptFooter: 'Thank you for your business!'
    };
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
    const companySettings = getCompanySettings();
    
    const now = new Date();
    const subtotal = total / (1 + companySettings.taxRate);
    const tax = total - subtotal;
    const discount = saleData.discount || 0;
    
    const receipt = {
      // Header
      companyName: companySettings.name,
      companyTagline: companySettings.tagline,
      companyLogo: companySettings.logo,
      companyAddress: companySettings.address,
      companyPhone: companySettings.phone,
      companyEmail: companySettings.email,
      companyWebsite: companySettings.website,
      businessHours: companySettings.businessHours,
      
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
      taxRate: (companySettings.taxRate * 100).toFixed(2),
      discount: discount,
      total: total,
      currency: companySettings.currency,
      
      // Payment
      paymentMethod: saleData.paymentMethod,
      amountPaid: saleData.amountPaid || total,
      change: (saleData.amountPaid || total) - total,
      
      // Footer
      footerMessage: companySettings.receiptFooter,
      
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

/**
 * ENTERPRISE POS FEATURES - Comparable to Micros/Symphony
 */

/**
 * Create Tables/Seating Management sheet
 */
function createTablesSheet(ss) {
  let sheet = ss.getSheetByName('Tables');
  if (!sheet) {
    sheet = ss.insertSheet('Tables');
  }
  
  sheet.clear();
  
  const headers = [
    'Table ID', 'Table Number', 'Section', 'Capacity', 'Status', 
    'Server ID', 'Check Number', 'Opened Time', 'Total Amount', 
    'Guest Count', 'Last Order Time', 'Special Notes'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#e91e63');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Create Menu Categories/Engineering sheet
 */
function createMenuCategoriesSheet(ss) {
  let sheet = ss.getSheetByName('MenuCategories');
  if (!sheet) {
    sheet = ss.insertSheet('MenuCategories');
  }
  
  sheet.clear();
  
  const headers = [
    'Category ID', 'Category Name', 'Display Order', 'Kitchen Station', 
    'Print To', 'Course Number', 'Active', 'Color Code', 'Icon', 
    'Preparation Time', 'Description'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#ff5722');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Create Kitchen Display System sheet
 */
function createKitchenOrdersSheet(ss) {
  let sheet = ss.getSheetByName('KitchenOrders');
  if (!sheet) {
    sheet = ss.insertSheet('KitchenOrders');
  }
  
  sheet.clear();
  
  const headers = [
    'Order ID', 'Table Number', 'Item Name', 'Quantity', 'Modifiers', 
    'Special Instructions', 'Station', 'Status', 'Order Time', 
    'Start Time', 'Complete Time', 'Priority', 'Server', 'Course'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#795548');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Create Inventory Adjustments sheet
 */
function createInventoryAdjustmentsSheet(ss) {
  let sheet = ss.getSheetByName('InventoryAdjustments');
  if (!sheet) {
    sheet = ss.insertSheet('InventoryAdjustments');
  }
  
  sheet.clear();
  
  const headers = [
    'Adjustment ID', 'Product ID', 'Product Name', 'Previous Stock', 
    'Adjustment Qty', 'New Stock', 'Reason Code', 'Reason Description', 
    'User', 'Date Time', 'Cost Impact', 'Reference Number'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#607d8b');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Create Promotions/Discounts sheet
 */
function createPromotionsSheet(ss) {
  let sheet = ss.getSheetByName('Promotions');
  if (!sheet) {
    sheet = ss.insertSheet('Promotions');
  }
  
  sheet.clear();
  
  const headers = [
    'Promo ID', 'Name', 'Type', 'Value', 'Min Purchase', 'Max Discount', 
    'Start Date', 'End Date', 'Start Time', 'End Time', 'Days Valid', 
    'Usage Limit', 'Used Count', 'Items/Categories', 'Active', 'Description'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Create Time Clock sheet for employee time tracking
 */
function createTimeClockSheet(ss) {
  let sheet = ss.getSheetByName('TimeClock');
  if (!sheet) {
    sheet = ss.insertSheet('TimeClock');
  }
  
  sheet.clear();
  
  const headers = [
    'Employee ID', 'Clock In', 'Clock Out', 'Break Start', 'Break End', 
    'Total Hours', 'Regular Hours', 'Overtime Hours', 'Break Minutes', 
    'Job Code', 'Department', 'Tips Declared', 'Hourly Rate'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#3f51b5');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Create Permissions Matrix sheet
 */
function createPermissionsSheet(ss) {
  let sheet = ss.getSheetByName('Permissions');
  if (!sheet) {
    sheet = ss.insertSheet('Permissions');
  }
  
  sheet.clear();
  
  const headers = [
    'Permission ID', 'Permission Name', 'Category', 'Description', 
    'Admin', 'Manager', 'Supervisor', 'Cashier', 'Server', 'Kitchen', 'Host'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#2196f3');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  
  // Add default permissions
  const defaultPermissions = [
    ['POS_SALES', 'Process Sales', 'POS', 'Ability to process sales transactions', true, true, true, true, true, false, false],
    ['POS_REFUNDS', 'Process Refunds', 'POS', 'Ability to process refunds and returns', true, true, true, false, false, false, false],
    ['POS_VOIDS', 'Void Transactions', 'POS', 'Ability to void transactions', true, true, true, false, false, false, false],
    ['POS_DISCOUNTS', 'Apply Discounts', 'POS', 'Ability to apply manual discounts', true, true, true, false, false, false, false],
    
    ['INVENTORY_VIEW', 'View Inventory', 'Inventory', 'View inventory levels and reports', true, true, true, true, false, false, false],
    ['INVENTORY_ADJUST', 'Adjust Inventory', 'Inventory', 'Adjust inventory quantities', true, true, false, false, false, false, false],
    ['INVENTORY_RECEIVE', 'Receive Inventory', 'Inventory', 'Receive new inventory shipments', true, true, true, false, false, false, false],
    ['INVENTORY_TRANSFER', 'Transfer Inventory', 'Inventory', 'Transfer inventory between locations', true, true, false, false, false, false, false],
    
    ['EMPLOYEE_VIEW', 'View Employees', 'Employee', 'View employee information', true, true, true, false, false, false, false],
    ['EMPLOYEE_MANAGE', 'Manage Employees', 'Employee', 'Add, edit, delete employees', true, true, false, false, false, false, false],
    ['EMPLOYEE_TIMECLOCK', 'Time Clock Access', 'Employee', 'Clock in/out and view time records', true, true, true, true, true, true, true],
    ['EMPLOYEE_PAYROLL', 'Payroll Access', 'Employee', 'View and manage payroll data', true, true, false, false, false, false, false],
    
    ['REPORTS_SALES', 'Sales Reports', 'Reports', 'View sales reports and analytics', true, true, true, false, false, false, false],
    ['REPORTS_INVENTORY', 'Inventory Reports', 'Reports', 'View inventory reports', true, true, true, false, false, false, false],
    ['REPORTS_EMPLOYEE', 'Employee Reports', 'Reports', 'View employee performance reports', true, true, true, false, false, false, false],
    ['REPORTS_FINANCIAL', 'Financial Reports', 'Reports', 'View financial and P&L reports', true, true, false, false, false, false, false],
    
    ['SETTINGS_SYSTEM', 'System Settings', 'Settings', 'Modify system settings', true, false, false, false, false, false, false],
    ['SETTINGS_COMPANY', 'Company Settings', 'Settings', 'Modify company information', true, true, false, false, false, false, false],
    ['SETTINGS_TAX', 'Tax Settings', 'Settings', 'Modify tax rates and rules', true, true, false, false, false, false, false],
    ['SETTINGS_PROMOTIONS', 'Promotion Settings', 'Settings', 'Create and manage promotions', true, true, false, false, false, false, false],
    
    ['TABLE_MANAGE', 'Table Management', 'Restaurant', 'Open, close, and manage tables', true, true, true, false, true, false, true],
    ['KITCHEN_DISPLAY', 'Kitchen Display', 'Restaurant', 'Access kitchen display system', true, true, true, false, false, true, false],
    ['MENU_MANAGE', 'Menu Management', 'Restaurant', 'Manage menu items and categories', true, true, false, false, false, false, false],
    ['FLOOR_PLAN', 'Floor Plan', 'Restaurant', 'View and modify floor plan', true, true, true, false, false, false, true],
    
    ['BACKUP_CREATE', 'Create Backups', 'System', 'Create system backups', true, false, false, false, false, false, false],
    ['LOGS_VIEW', 'View System Logs', 'System', 'View system activity logs', true, true, false, false, false, false, false],
    ['CASH_DRAWER', 'Cash Drawer', 'Cash', 'Open cash drawer', true, true, true, true, false, false, false],
    ['CASH_COUNT', 'Cash Count', 'Cash', 'Perform cash counts and reconciliation', true, true, true, false, false, false, false]
  ];
  
  defaultPermissions.forEach((permission, index) => {
    sheet.getRange(index + 2, 1, 1, permission.length).setValues([permission]);
  });
  
  return sheet;
}

/**
 * Create Roles sheet for custom role management
 */
function createRolesSheet(ss) {
  let sheet = ss.getSheetByName('Roles');
  if (!sheet) {
    sheet = ss.insertSheet('Roles');
  }
  
  sheet.clear();
  
  const headers = [
    'Role ID', 'Role Name', 'Description', 'Hourly Rate Min', 'Hourly Rate Max', 
    'Can Override Prices', 'Max Discount %', 'Active', 'Created Date', 'Color Code'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4caf50');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  
  // Add default roles
  const defaultRoles = [
    ['ADMIN', 'Administrator', 'Full system access', 0, 0, true, 100, true, new Date(), '#f44336'],
    ['MANAGER', 'Manager', 'Management level access', 20, 35, true, 50, true, new Date(), '#ff9800'],
    ['SUPERVISOR', 'Supervisor', 'Supervisory access', 15, 25, true, 25, true, new Date(), '#2196f3'],
    ['CASHIER', 'Cashier', 'POS and basic functions', 12, 18, false, 10, true, new Date(), '#4caf50'],
    ['SERVER', 'Server', 'Table service and POS', 8, 15, false, 5, true, new Date(), '#9c27b0'],
    ['KITCHEN', 'Kitchen Staff', 'Kitchen display and prep', 14, 22, false, 0, true, new Date(), '#795548'],
    ['HOST', 'Host/Hostess', 'Seating and basic functions', 10, 16, false, 0, true, new Date(), '#607d8b']
  ];
  
  defaultRoles.forEach((role, index) => {
    sheet.getRange(index + 2, 1, 1, role.length).setValues([role]);
  });
  
  return sheet;
}

/**
 * Get user permissions based on role
 */
function getUserPermissions(userRole) {
  try {
    const permissionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Permissions');
    const data = permissionsSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find the role column
    const roleColumnIndex = headers.findIndex(header => header.toLowerCase() === userRole.toLowerCase());
    if (roleColumnIndex === -1) {
      console.error('Role not found:', userRole);
      return [];
    }
    
    const permissions = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][roleColumnIndex] === true) {
        permissions.push({
          id: data[i][0],
          name: data[i][1],
          category: data[i][2],
          description: data[i][3]
        });
      }
    }
    
    return permissions;
    
  } catch (error) {
    console.error('Get user permissions error:', error);
    return [];
  }
}

/**
 * Check if user has specific permission
 */
function hasPermission(userRole, permissionID) {
  try {
    const permissions = getUserPermissions(userRole);
    return permissions.some(permission => permission.id === permissionID);
  } catch (error) {
    console.error('Check permission error:', error);
    return false;
  }
}

/**
 * Update permission for a role
 */
function updateRolePermission(role, permissionID, hasAccess) {
  try {
    const permissionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Permissions');
    const data = permissionsSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find the role column and permission row
    const roleColumnIndex = headers.findIndex(header => header.toLowerCase() === role.toLowerCase());
    if (roleColumnIndex === -1) {
      return { success: false, error: 'Role not found' };
    }
    
    let permissionRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === permissionID) {
        permissionRowIndex = i;
        break;
      }
    }
    
    if (permissionRowIndex === -1) {
      return { success: false, error: 'Permission not found' };
    }
    
    // Update the permission
    permissionsSheet.getRange(permissionRowIndex + 1, roleColumnIndex + 1).setValue(hasAccess);
    
    logAction('Permission Updated', 'Admin', `${role} - ${permissionID}: ${hasAccess}`, currentSessionId);
    
    return { success: true };
    
  } catch (error) {
    console.error('Update role permission error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Get all permissions matrix for admin panel
 */
function getPermissionsMatrix() {
  try {
    const permissionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Permissions');
    const data = permissionsSheet.getDataRange().getValues();
    
    if (data.length === 0) {
      return { success: false, error: 'No permissions data found' };
    }
    
    const headers = data[0];
    const roleColumns = headers.slice(4); // Skip first 4 columns (ID, Name, Category, Description)
    
    const permissions = [];
    for (let i = 1; i < data.length; i++) {
      const permission = {
        id: data[i][0],
        name: data[i][1],
        category: data[i][2],
        description: data[i][3],
        roles: {}
      };
      
      // Map role permissions
      roleColumns.forEach((role, index) => {
        permission.roles[role] = data[i][4 + index];
      });
      
      permissions.push(permission);
    }
    
    return {
      success: true,
      permissions: permissions,
      roles: roleColumns
    };
    
  } catch (error) {
    console.error('Get permissions matrix error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Get role information
 */
function getRoleInfo(roleID) {
  try {
    const rolesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roles');
    const data = rolesSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === roleID) {
        return {
          success: true,
          role: {
            id: data[i][0],
            name: data[i][1],
            description: data[i][2],
            hourlyRateMin: data[i][3],
            hourlyRateMax: data[i][4],
            canOverridePrices: data[i][5],
            maxDiscountPercent: data[i][6],
            active: data[i][7],
            createdDate: data[i][8],
            colorCode: data[i][9]
          }
        };
      }
    }
    
    return { success: false, error: 'Role not found' };
    
  } catch (error) {
    console.error('Get role info error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * ENTERPRISE BUSINESS LOGIC FUNCTIONS
 */

/**
 * Table Management Functions
 */
function openTable(tableNumber, serverID, guestCount, section = 'Main') {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tables');
    const checkNumber = generateCheckNumber();
    const now = new Date();
    
    const tableData = [
      generateTableID(),
      tableNumber,
      section,
      guestCount,
      'OCCUPIED',
      serverID,
      checkNumber,
      now,
      0,
      guestCount,
      now,
      ''
    ];
    
    sheet.appendRow(tableData);
    
    logAction('Table Opened', serverID, `Table ${tableNumber} opened for ${guestCount} guests`, currentSessionId);
    
    return {
      success: true,
      checkNumber: checkNumber,
      tableID: tableData[0]
    };
    
  } catch (error) {
    console.error('Open table error:', error);
    return { success: false, error: error.toString() };
  }
}

function closeTable(tableID, serverID) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tables');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === tableID) {
        sheet.getRange(i + 1, 5).setValue('AVAILABLE'); // Status
        sheet.getRange(i + 1, 6).setValue(''); // Server ID
        sheet.getRange(i + 1, 7).setValue(''); // Check Number
        sheet.getRange(i + 1, 8).setValue(''); // Opened Time
        sheet.getRange(i + 1, 9).setValue(0); // Total Amount
        sheet.getRange(i + 1, 12).setValue(''); // Special Notes
        
        logAction('Table Closed', serverID, `Table ${data[i][1]} closed`, currentSessionId);
        
        return { success: true };
      }
    }
    
    return { success: false, error: 'Table not found' };
    
  } catch (error) {
    console.error('Close table error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Kitchen Display System Functions
 */
function sendOrderToKitchen(orderData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KitchenOrders');
    const now = new Date();
    
    orderData.items.forEach(item => {
      const kitchenOrder = [
        generateOrderID(),
        orderData.tableNumber,
        item.name,
        item.quantity,
        item.modifiers || '',
        item.specialInstructions || '',
        item.kitchenStation || 'HOT',
        'PENDING',
        now,
        '',
        '',
        item.priority || 'NORMAL',
        orderData.serverID,
        item.course || 1
      ];
      
      sheet.appendRow(kitchenOrder);
    });
    
    logAction('Order Sent to Kitchen', orderData.serverID, `${orderData.items.length} items sent to kitchen for table ${orderData.tableNumber}`, currentSessionId);
    
    return { success: true };
    
  } catch (error) {
    console.error('Send order to kitchen error:', error);
    return { success: false, error: error.toString() };
  }
}

function updateKitchenOrderStatus(orderID, status, station) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KitchenOrders');
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === orderID) {
        sheet.getRange(i + 1, 8).setValue(status); // Status
        
        if (status === 'COOKING') {
          sheet.getRange(i + 1, 10).setValue(now); // Start Time
        } else if (status === 'READY') {
          sheet.getRange(i + 1, 11).setValue(now); // Complete Time
        }
        
        logAction('Kitchen Order Update', station, `Order ${orderID} status changed to ${status}`, currentSessionId);
        
        return { success: true };
      }
    }
    
    return { success: false, error: 'Order not found' };
    
  } catch (error) {
    console.error('Update kitchen order error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Advanced Inventory Management
 */
function adjustInventory(productID, adjustmentQty, reasonCode, reasonDescription, userID) {
  try {
    const productsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PRODUCTS);
    const adjustmentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('InventoryAdjustments');
    const productsData = productsSheet.getDataRange().getValues();
    
    let productFound = false;
    let previousStock = 0;
    let productName = '';
    let productCost = 0;
    
    // Find product and update stock
    for (let i = 1; i < productsData.length; i++) {
      if (productsData[i][0] === productID) {
        previousStock = productsData[i][3];
        productName = productsData[i][1];
        productCost = productsData[i][2];
        
        const newStock = previousStock + adjustmentQty;
        productsSheet.getRange(i + 1, 4).setValue(newStock); // Update stock
        
        productFound = true;
        break;
      }
    }
    
    if (!productFound) {
      return { success: false, error: 'Product not found' };
    }
    
    // Record adjustment
    const costImpact = adjustmentQty * productCost;
    const adjustmentData = [
      generateAdjustmentID(),
      productID,
      productName,
      previousStock,
      adjustmentQty,
      previousStock + adjustmentQty,
      reasonCode,
      reasonDescription,
      userID,
      new Date(),
      costImpact,
      generateReferenceNumber()
    ];
    
    adjustmentsSheet.appendRow(adjustmentData);
    
    logAction('Inventory Adjustment', userID, `${productName}: ${adjustmentQty} units (${reasonCode})`, currentSessionId, costImpact);
    
    return {
      success: true,
      newStock: previousStock + adjustmentQty,
      costImpact: costImpact
    };
    
  } catch (error) {
    console.error('Inventory adjustment error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Employee Time Clock Functions
 */
function clockIn(employeeID, jobCode = 'SERVER', department = 'FRONT') {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TimeClock');
    const now = new Date();
    
    // Check if already clocked in
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === employeeID && !data[i][2]) { // No clock out time
        return { success: false, error: 'Employee already clocked in' };
      }
    }
    
    const timeEntry = [
      employeeID,
      now,
      '', // Clock out (empty)
      '', // Break start (empty)
      '', // Break end (empty)
      0,  // Total hours
      0,  // Regular hours
      0,  // Overtime hours
      0,  // Break minutes
      jobCode,
      department,
      0,  // Tips declared
      getEmployeeHourlyRate(employeeID)
    ];
    
    sheet.appendRow(timeEntry);
    
    logAction('Clock In', employeeID, `Clocked in as ${jobCode} in ${department}`, currentSessionId);
    
    return { success: true, clockInTime: now };
    
  } catch (error) {
    console.error('Clock in error:', error);
    return { success: false, error: error.toString() };
  }
}

function clockOut(employeeID, tipsDeclared = 0) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TimeClock');
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    
    // Find open clock in record
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === employeeID && !data[i][2]) { // No clock out time
        const clockInTime = new Date(data[i][1]);
        const totalMinutes = Math.round((now - clockInTime) / (1000 * 60));
        const totalHours = totalMinutes / 60;
        const regularHours = Math.min(totalHours, 8);
        const overtimeHours = Math.max(totalHours - 8, 0);
        
        sheet.getRange(i + 1, 3).setValue(now); // Clock out
        sheet.getRange(i + 1, 6).setValue(totalHours); // Total hours
        sheet.getRange(i + 1, 7).setValue(regularHours); // Regular hours
        sheet.getRange(i + 1, 8).setValue(overtimeHours); // Overtime hours
        sheet.getRange(i + 1, 12).setValue(tipsDeclared); // Tips declared
        
        logAction('Clock Out', employeeID, `Worked ${totalHours.toFixed(2)} hours, Tips: $${tipsDeclared}`, currentSessionId);
        
        return {
          success: true,
          totalHours: totalHours,
          regularHours: regularHours,
          overtimeHours: overtimeHours,
          tipsDeclared: tipsDeclared
        };
      }
    }
    
    return { success: false, error: 'No active clock in found' };
    
  } catch (error) {
    console.error('Clock out error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Promotions and Discounts Engine
 */
function calculatePromotionalDiscount(items, customerID = null, promoCode = null) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Promotions');
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const currentTime = now.getTime();
    
    let bestDiscount = 0;
    let appliedPromo = null;
    const subtotal = items.reduce((sum, item) => sum + (item.price * item.quantity), 0);
    
    for (let i = 1; i < data.length; i++) {
      const [promoID, name, type, value, minPurchase, maxDiscount, startDate, endDate, startTime, endTime, daysValid, usageLimit, usedCount, itemsCategories, active] = data[i];
      
      if (!active) continue;
      
      // Check date validity
      const start = new Date(startDate);
      const end = new Date(endDate);
      if (currentTime < start.getTime() || currentTime > end.getTime()) continue;
      
      // Check minimum purchase
      if (subtotal < minPurchase) continue;
      
      // Check usage limit
      if (usageLimit > 0 && usedCount >= usageLimit) continue;
      
      // Calculate discount
      let discount = 0;
      if (type === 'PERCENTAGE') {
        discount = subtotal * (value / 100);
      } else if (type === 'FIXED') {
        discount = value;
      } else if (type === 'BUY_X_GET_Y') {
        // Implement buy X get Y logic
        discount = calculateBuyXGetYDiscount(items, value);
      }
      
      // Apply maximum discount limit
      if (maxDiscount > 0) {
        discount = Math.min(discount, maxDiscount);
      }
      
      if (discount > bestDiscount) {
        bestDiscount = discount;
        appliedPromo = {
          id: promoID,
          name: name,
          type: type,
          discount: discount
        };
      }
    }
    
    return {
      success: true,
      discount: bestDiscount,
      promotion: appliedPromo
    };
    
  } catch (error) {
    console.error('Calculate promotional discount error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Helper Functions for Enterprise Features
 */
function generateTableID() {
  return 'TBL' + new Date().getTime() + Math.floor(Math.random() * 100);
}

function generateCheckNumber() {
  return 'CHK' + new Date().getTime().toString().slice(-8);
}

function generateOrderID() {
  return 'ORD' + new Date().getTime() + Math.floor(Math.random() * 100);
}

function generateAdjustmentID() {
  return 'ADJ' + new Date().getTime() + Math.floor(Math.random() * 100);
}

function generateReferenceNumber() {
  return 'REF' + new Date().getTime().toString().slice(-10);
}

function getEmployeeHourlyRate(employeeID) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === employeeID) {
        return data[i][9] || 15.00; // Default hourly rate
      }
    }
    
    return 15.00; // Default rate
  } catch (error) {
    return 15.00;
  }
}

function calculateBuyXGetYDiscount(items, value) {
  // Simplified buy X get Y calculation
  // In real implementation, this would be more complex
  const eligibleItems = items.filter(item => item.category === 'PROMO_ELIGIBLE');
  if (eligibleItems.length >= value) {
    const cheapestItem = eligibleItems.sort((a, b) => a.price - b.price)[0];
    return cheapestItem.price;
  }
  return 0;
}

/**
 * ADVANCED POS FEATURES
 */

/**
 * Create Customers sheet for customer management
 */
function createCustomersSheet(ss) {
  let sheet = ss.getSheetByName('Customers');
  if (!sheet) {
    sheet = ss.insertSheet('Customers');
  }
  
  sheet.clear();
  
  const headers = [
    'Customer ID', 'First Name', 'Last Name', 'Email', 'Phone', 'Date of Birth',
    'Address', 'City', 'State', 'ZIP', 'Country', 'Loyalty Points', 'Total Spent',
    'Visit Count', 'Last Visit', 'Preferred Payment', 'Notes', 'Tags', 
    'Discount Level', 'Created Date', 'Status', 'Marketing Consent'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#8e24aa');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  
  return sheet;
}

/**
 * Create Split Payments sheet
 */
function createSplitPaymentsSheet(ss) {
  let sheet = ss.getSheetByName('SplitPayments');
  if (!sheet) {
    sheet = ss.insertSheet('SplitPayments');
  }
  
  sheet.clear();
  
  const headers = [
    'Transaction ID', 'Payment ID', 'Payment Method', 'Amount', 'Currency',
    'Card Last 4', 'Auth Code', 'Reference Number', 'Status', 'Timestamp',
    'Gateway', 'Fee Amount', 'Net Amount', 'Customer ID'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#00695c');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  
  return sheet;
}

/**
 * Create Loyalty Program sheet
 */
function createLoyaltyProgramSheet(ss) {
  let sheet = ss.getSheetByName('LoyaltyProgram');
  if (!sheet) {
    sheet = ss.insertSheet('LoyaltyProgram');
  }
  
  sheet.clear();
  
  const headers = [
    'Program ID', 'Program Name', 'Points Per Dollar', 'Reward Threshold', 'Reward Value',
    'Bonus Points Events', 'Expiry Days', 'Tier Requirements', 'Tier Benefits',
    'Active', 'Start Date', 'End Date', 'Description'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1565c0');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  
  // Add default loyalty program
  const defaultProgram = [
    'BASIC_LOYALTY', 'Basic Loyalty Program', 1, 100, 5,
    'Double points on weekends', 365, 'Bronze: 0, Silver: 500, Gold: 1000', 
    'Bronze: 5% discount, Silver: 10% discount, Gold: 15% discount',
    true, new Date(), null, 'Standard loyalty program with tiered benefits'
  ];
  sheet.appendRow(defaultProgram);
  
  return sheet;
}

/**
 * Create Advanced Discounts sheet
 */
function createAdvancedDiscountsSheet(ss) {
  let sheet = ss.getSheetByName('AdvancedDiscounts');
  if (!sheet) {
    sheet = ss.insertSheet('AdvancedDiscounts');
  }
  
  sheet.clear();
  
  const headers = [
    'Discount ID', 'Name', 'Type', 'Value', 'Min Quantity', 'Max Discount',
    'Applies To', 'Customer Groups', 'Time Restrictions', 'Date Range',
    'Usage Limit', 'Used Count', 'Stackable', 'Auto Apply', 'Priority',
    'Conditions', 'Actions', 'Active', 'Created By', 'Created Date'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#d32f2f');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  
  return sheet;
}

/**
 * Customer Management Functions
 */
function addCustomer(customerData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    const customerId = generateCustomerId();
    
    const row = [
      customerId,
      customerData.firstName,
      customerData.lastName,
      customerData.email,
      customerData.phone,
      customerData.dateOfBirth || '',
      customerData.address || '',
      customerData.city || '',
      customerData.state || '',
      customerData.zip || '',
      customerData.country || '',
      0, // Loyalty points
      0, // Total spent
      0, // Visit count
      '', // Last visit
      customerData.preferredPayment || '',
      customerData.notes || '',
      customerData.tags || '',
      customerData.discountLevel || 'NONE',
      new Date(),
      'ACTIVE',
      customerData.marketingConsent || false
    ];
    
    sheet.appendRow(row);
    
    logAction('Customer Added', customerData.addedBy || 'System', `Customer added: ${customerData.firstName} ${customerData.lastName}`, currentSessionId);
    
    return {
      success: true,
      customerId: customerId
    };
    
  } catch (error) {
    console.error('Add customer error:', error);
    return { success: false, error: error.toString() };
  }
}

function updateCustomerLoyalty(customerId, purchaseAmount) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === customerId) {
        const currentPoints = data[i][11] || 0;
        const currentSpent = data[i][12] || 0;
        const currentVisits = data[i][13] || 0;
        
        // Calculate new points (1 point per dollar)
        const newPoints = currentPoints + Math.floor(purchaseAmount);
        const newSpent = currentSpent + purchaseAmount;
        const newVisits = currentVisits + 1;
        
        // Update customer record
        sheet.getRange(i + 1, 12).setValue(newPoints); // Loyalty points
        sheet.getRange(i + 1, 13).setValue(newSpent); // Total spent
        sheet.getRange(i + 1, 14).setValue(newVisits); // Visit count
        sheet.getRange(i + 1, 15).setValue(new Date()); // Last visit
        
        // Check for tier upgrades
        const newTier = calculateLoyaltyTier(newPoints);
        
        return {
          success: true,
          newPoints: newPoints,
          newTier: newTier,
          totalSpent: newSpent
        };
      }
    }
    
    return { success: false, error: 'Customer not found' };
    
  } catch (error) {
    console.error('Update customer loyalty error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Split Payment Processing
 */
function processSplitPayment(transactionId, payments) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SplitPayments');
    const results = [];
    
    payments.forEach(payment => {
      const paymentId = generatePaymentId();
      const now = new Date();
      
      const paymentRow = [
        transactionId,
        paymentId,
        payment.method,
        payment.amount,
        payment.currency || 'USD',
        payment.cardLast4 || '',
        payment.authCode || '',
        payment.referenceNumber || '',
        'COMPLETED',
        now,
        payment.gateway || 'INTERNAL',
        payment.feeAmount || 0,
        payment.amount - (payment.feeAmount || 0),
        payment.customerId || ''
      ];
      
      sheet.appendRow(paymentRow);
      
      results.push({
        paymentId: paymentId,
        method: payment.method,
        amount: payment.amount,
        status: 'COMPLETED'
      });
    });
    
    logAction('Split Payment', 'System', `Split payment processed: ${payments.length} methods, Total: $${payments.reduce((sum, p) => sum + p.amount, 0).toFixed(2)}`, currentSessionId);
    
    return {
      success: true,
      payments: results
    };
    
  } catch (error) {
    console.error('Process split payment error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Advanced Discount Engine
 */
function calculateAdvancedDiscounts(items, customer, promoCode) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AdvancedDiscounts');
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    
    let applicableDiscounts = [];
    let totalDiscount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const [discountId, name, type, value, minQty, maxDiscount, appliesTo, customerGroups, timeRestrictions, dateRange, usageLimit, usedCount, stackable, autoApply, priority, conditions, actions, active] = data[i];
      
      if (!active) continue;
      
      // Check if discount applies
      if (autoApply || promoCode === discountId) {
        let discount = 0;
        
        switch (type) {
          case 'PERCENTAGE':
            discount = calculatePercentageDiscount(items, value, appliesTo);
            break;
          case 'FIXED':
            discount = Math.min(value, items.reduce((sum, item) => sum + (item.price * item.quantity), 0));
            break;
          case 'BUY_X_GET_Y':
            discount = calculateBuyXGetYDiscount(items, conditions);
            break;
          case 'LOYALTY':
            if (customer && customer.loyaltyPoints) {
              discount = calculateLoyaltyDiscount(customer, value);
            }
            break;
        }
        
        if (maxDiscount > 0) {
          discount = Math.min(discount, maxDiscount);
        }
        
        if (discount > 0) {
          applicableDiscounts.push({
            id: discountId,
            name: name,
            type: type,
            discount: discount,
            priority: priority || 0
          });
        }
      }
    }
    
    // Sort by priority and apply
    applicableDiscounts.sort((a, b) => b.priority - a.priority);
    
    return {
      success: true,
      discounts: applicableDiscounts,
      totalDiscount: applicableDiscounts.reduce((sum, d) => sum + d.discount, 0)
    };
    
  } catch (error) {
    console.error('Calculate advanced discounts error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Helper Functions
 */
function generateCustomerId() {
  return 'CUST' + new Date().getTime() + Math.floor(Math.random() * 100);
}

function generatePaymentId() {
  return 'PAY' + new Date().getTime() + Math.floor(Math.random() * 100);
}

function calculateLoyaltyTier(points) {
  if (points >= 1000) return 'GOLD';
  if (points >= 500) return 'SILVER';
  return 'BRONZE';
}

function calculatePercentageDiscount(items, percentage, appliesTo) {
  let applicableAmount = 0;
  
  items.forEach(item => {
    if (appliesTo === 'ALL' || appliesTo.includes(item.category)) {
      applicableAmount += item.price * item.quantity;
    }
  });
  
  return applicableAmount * (percentage / 100);
}

function calculateBuyXGetYDiscount(items, conditions) {
  // Simplified implementation
  // In real scenario, this would parse complex conditions
  return 0;
}

function calculateLoyaltyDiscount(customer, discountRate) {
  const tier = calculateLoyaltyTier(customer.loyaltyPoints);
  
  switch (tier) {
    case 'GOLD': return discountRate * 1.5;
    case 'SILVER': return discountRate * 1.2;
    default: return discountRate;
  }
}

/**
 * CUSTOMIZATION & SETTINGS MANAGEMENT
 */

/**
 * Save branding settings to Google Sheets
 */
function saveBrandingSettings(brandingData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
    
    // Update or add branding settings
    const settingsToUpdate = [
      ['COMPANY_NAME', brandingData.companyName],
      ['COMPANY_TAGLINE', brandingData.tagline],
      ['COMPANY_LOGO', brandingData.logo],
      ['PRIMARY_COLOR', brandingData.primaryColor]
    ];
    
    settingsToUpdate.forEach(([key, value]) => {
      updateSetting(key, value);
    });
    
    logAction('Branding Updated', currentUser?.username || 'Admin', 'Branding settings updated', currentSessionId);
    
    return { success: true };
    
  } catch (error) {
    console.error('Save branding settings error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Save POS configuration
 */
function savePOSConfiguration(configData) {
  try {
    const settingsToUpdate = [
      ['TAX_RATE', configData.taxRate],
      ['CURRENCY', configData.currency],
      ['RECEIPT_FOOTER', configData.receiptFooter]
    ];
    
    settingsToUpdate.forEach(([key, value]) => {
      updateSetting(key, value);
    });
    
    logAction('Configuration Updated', currentUser?.username || 'Admin', 'POS configuration updated', currentSessionId);
    
    return { success: true };
    
  } catch (error) {
    console.error('Save POS configuration error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Save feature toggles
 */
function saveFeatureToggles(features) {
  try {
    Object.keys(features).forEach(feature => {
      const settingKey = `FEATURE_${feature.toUpperCase()}`;
      updateSetting(settingKey, features[feature]);
    });
    
    logAction('Features Updated', currentUser?.username || 'Admin', 'Feature toggles updated', currentSessionId);
    
    return { success: true };
    
  } catch (error) {
    console.error('Save feature toggles error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Get all customization settings
 */
function getCustomizationSettings() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
    const data = sheet.getDataRange().getValues();
    
    const settings = {
      branding: {},
      config: {},
      features: {}
    };
    
    for (let i = 1; i < data.length; i++) {
      const [key, value] = data[i];
      
      // Categorize settings
      if (key.startsWith('COMPANY_') || key === 'PRIMARY_COLOR') {
        const brandingKey = key.replace('COMPANY_', '').toLowerCase();
        if (brandingKey === 'name') settings.branding.companyName = value;
        else if (brandingKey === 'tagline') settings.branding.tagline = value;
        else if (brandingKey === 'logo') settings.branding.logo = value;
        else if (key === 'PRIMARY_COLOR') settings.branding.primaryColor = value;
      } else if (['TAX_RATE', 'CURRENCY', 'RECEIPT_FOOTER'].includes(key)) {
        const configKey = key.toLowerCase().replace('_', '');
        if (key === 'TAX_RATE') settings.config.taxRate = value;
        else if (key === 'CURRENCY') settings.config.currency = value;
        else if (key === 'RECEIPT_FOOTER') settings.config.receiptFooter = value;
      } else if (key.startsWith('FEATURE_')) {
        const featureKey = key.replace('FEATURE_', '').toLowerCase();
        settings.features[featureKey] = value;
      }
    }
    
    return {
      success: true,
      ...settings
    };
    
  } catch (error) {
    console.error('Get customization settings error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Get integrations status (placeholder)
 */
function getIntegrationsStatus() {
  try {
    // This would normally check actual integration statuses
    // For demo purposes, showing some connected/disconnected states
    const integrations = {
      stripe: { connected: false, lastSync: null },
      square: { connected: false, lastSync: null },
      quickbooks: { connected: true, lastSync: new Date() },
      mailchimp: { connected: false, lastSync: null },
      twilio: { connected: false, lastSync: null },
      analytics: { connected: true, lastSync: new Date() }
    };
    
    return {
      success: true,
      data: integrations
    };
    
  } catch (error) {
    console.error('Get integrations status error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Helper function to update a setting
 */
function updateSetting(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      found = true;
      break;
    }
  }
  
  if (!found) {
    sheet.appendRow([key, value]);
  }
}

/**
 * ENHANCED CUSTOMER FUNCTIONS
 */

/**
 * Search customers by various criteria
 */
function searchCustomers(searchTerm, searchType = 'all') {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    const data = sheet.getDataRange().getValues();
    const customers = [];
    
    for (let i = 1; i < data.length; i++) {
      const [customerId, firstName, lastName, email, phone, , address, city, state, zip, country, loyaltyPoints, totalSpent, visitCount, lastVisit, preferredPayment, notes, tags, discountLevel, createdDate, status] = data[i];
      
      let matches = false;
      const searchLower = searchTerm.toLowerCase();
      
      switch (searchType) {
        case 'name':
          matches = firstName.toLowerCase().includes(searchLower) || lastName.toLowerCase().includes(searchLower);
          break;
        case 'email':
          matches = email.toLowerCase().includes(searchLower);
          break;
        case 'phone':
          matches = phone.includes(searchTerm);
          break;
        case 'id':
          matches = customerId.toLowerCase().includes(searchLower);
          break;
        default:
          matches = firstName.toLowerCase().includes(searchLower) || 
                   lastName.toLowerCase().includes(searchLower) ||
                   email.toLowerCase().includes(searchLower) ||
                   phone.includes(searchTerm) ||
                   customerId.toLowerCase().includes(searchLower);
      }
      
      if (matches && status === 'ACTIVE') {
        customers.push({
          customerId,
          firstName,
          lastName,
          email,
          phone,
          address,
          city,
          state,
          zip,
          country,
          loyaltyPoints,
          totalSpent,
          visitCount,
          lastVisit,
          preferredPayment,
          notes,
          tags,
          discountLevel,
          createdDate,
          status
        });
      }
    }
    
    return {
      success: true,
      customers: customers
    };
    
  } catch (error) {
    console.error('Search customers error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Get customer details by ID
 */
function getCustomerById(customerId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === customerId) {
        const [, firstName, lastName, email, phone, dateOfBirth, address, city, state, zip, country, loyaltyPoints, totalSpent, visitCount, lastVisit, preferredPayment, notes, tags, discountLevel, createdDate, status] = data[i];
        
        return {
          success: true,
          customer: {
            customerId,
            firstName,
            lastName,
            email,
            phone,
            dateOfBirth,
            address,
            city,
            state,
            zip,
            country,
            loyaltyPoints,
            totalSpent,
            visitCount,
            lastVisit,
            preferredPayment,
            notes,
            tags,
            discountLevel,
            createdDate,
            status,
            tier: calculateLoyaltyTier(loyaltyPoints)
          }
        };
      }
    }
    
    return { success: false, error: 'Customer not found' };
    
  } catch (error) {
    console.error('Get customer by ID error:', error);
    return { success: false, error: error.toString() };
  }
}