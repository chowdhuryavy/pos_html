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
          
          logAction('User Login', username, `User ${name} logged in successfully using ${authType.toUpperCase()}`);
          
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