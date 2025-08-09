# 🚀 **POS System Setup Guide**

## **Complete Setup Instructions for Google Sheets + Apps Script**

### **📋 Prerequisites**
- Google account with access to Google Sheets and Apps Script
- Basic understanding of copy/paste operations
- 15-20 minutes setup time

---

## **🔧 Step 1: Create Google Spreadsheet**

1. **Go to [Google Sheets](https://sheets.google.com)**
2. **Click "+" to create a new blank spreadsheet**
3. **Rename it to "POS System Database"**
4. **Copy the URL** - you'll need it later

---

## **⚙️ Step 2: Set Up Google Apps Script**

### **2.1 Open Apps Script**
1. In your Google Sheet, go to **Extensions → Apps Script**
2. Delete the default `myFunction()` code
3. You should see an empty `Code.gs` file

### **2.2 Copy Backend Code**
1. **Copy ALL content from `/workspace/Code.gs`**
2. **Paste it into the Apps Script `Code.gs` file**
3. **Save** (Ctrl+S) and name the project "POS System"

### **2.3 Create HTML Files**
Create these files by clicking the **"+"** button next to "Files":

#### **File 1: index.html**
1. Click **"+" → HTML file**
2. Name it `index`
3. **Copy ALL content from `/workspace/index.html`**
4. **Paste and save**

#### **File 2: pos-script.html**
1. Click **"+" → HTML file**
2. Name it `pos-script`
3. **Copy ALL content from `/workspace/pos-script.html`**
4. **Paste and save**

#### **File 3: charts.html**
1. Click **"+" → HTML file**
2. Name it `charts`
3. **Copy ALL content from `/workspace/charts.html`**
4. **Paste and save**

---

## **🗃️ Step 3: Initialize Database**

### **3.1 Run Initialization**
1. **In Apps Script, make sure `Code.gs` is selected**
2. **Find the function dropdown (next to Debug)**
3. **Select `initializePOSSystem`**
4. **Click "Run" (▶️ button)**
5. **Authorize permissions when prompted**

### **3.2 Verify Setup**
1. **Select `testSystemSetup` from function dropdown**
2. **Click "Run" (▶️ button)**
3. **Check the output in the console**

**Expected Success Output:**
```
✅ SUCCESS: System setup complete!

📊 Database Summary:
- 3 employees created
- 5 products added
- All sheets configured
- Admin authentication working

🚀 Ready to deploy web app!
```

---

## **🌐 Step 4: Deploy Web Application**

### **4.1 Create Deployment**
1. **Click "Deploy" → "New deployment"**
2. **Click the gear icon ⚙️ next to "Type"**
3. **Select "Web app"**

### **4.2 Configure Deployment**
- **Description:** "POS System v1.0"
- **Execute as:** "Me"
- **Who has access:** "Anyone" (or "Anyone with Google account" for security)

### **4.3 Deploy**
1. **Click "Deploy"**
2. **Copy the Web app URL** - this is your POS system URL!
3. **Click "Done"**

---

## **🎯 Step 5: Test Your POS System**

### **5.1 Open Your POS System**
1. **Open the Web app URL in a new browser tab**
2. **You should see the beautiful login page**

### **5.2 Login with Default Credentials**

**Option 1 - PIN Authentication:**
- Username: `admin`
- PIN: `1234`

**Option 2 - Password Authentication:**
- Username: `admin`
- Password: `admin123`

**Additional Test Users:**
- Manager: `john` / PIN: `5678` / Password: `john123`
- Cashier: `sarah` / PIN: `9999` / Password: `sarah123`

---

## **📊 Step 6: Verify Database Structure**

After successful setup, your Google Spreadsheet should contain these sheets:

### **📦 Products Sheet**
| ID | Name | Price | Cost | Stock | Category | Barcode | Active |
|----|------|-------|------|-------|----------|---------|---------|
| PRD001 | Coffee | 3.50 | 1.20 | 50 | Beverages | 1234567890123 | TRUE |
| PRD002 | Sandwich | 8.99 | 4.50 | 25 | Food | 2345678901234 | TRUE |

### **👥 Employees Sheet**
| ID | Name | Role | Username | PIN Hash | Active | Created | Last Login | Password Hash |
|----|------|------|----------|----------|---------|---------|------------|---------------|
| EMP001 | Admin User | Admin | admin | [hashed] | TRUE | [date] | [null] | [hashed] |

### **💰 Sales Sheet**
| Sale ID | Date | Product ID | Product Name | Quantity | Unit Price | Total | Tax | Discount | Employee | Payment Method |
|---------|------|------------|--------------|----------|------------|-------|-----|----------|----------|----------------|
| (Empty initially) |

### **⚙️ Settings Sheet**
| Key | Value | Description |
|-----|-------|-------------|
| TAX_RATE | 0.0875 | Default tax rate (8.75%) |
| CURRENCY | USD | Currency symbol |
| COMPANY_NAME | RetailPro | Company name for branding |

### **📋 Logs Sheet**
| Timestamp | Action | User | Details |
|-----------|--------|------|---------|
| [datetime] | System Initialization | System | POS system initialized successfully |

---

## **🔧 Troubleshooting**

### **❌ Common Issues and Solutions**

#### **Issue: "Function not found" error**
- **Solution:** Make sure you copied ALL content from `Code.gs` completely
- Re-run `initializePOSSystem()`

#### **Issue: "Permission denied" error**
- **Solution:** 
  1. Go to Apps Script dashboard
  2. Click on your project
  3. Click "Review permissions"
  4. Authorize all required permissions

#### **Issue: Blank sheets created**
- **Solution:** Run `resetPOSSystem()` to clean and reinitialize

#### **Issue: Web app shows "Script function not found"**
- **Solution:** 
  1. Make sure all HTML files are created correctly
  2. Redeploy the web app
  3. Clear browser cache

#### **Issue: Login not working**
- **Solution:** 
  1. Run `testSystemSetup()` to verify database
  2. Check employee data in Employees sheet
  3. Verify authentication function in Apps Script

---

## **🚀 Advanced Configuration**

### **Company Branding**
Edit the **Settings** sheet to customize:
- `COMPANY_NAME` - Your business name
- `COMPANY_TAGLINE` - Business description
- `TAX_RATE` - Your local tax rate
- `CURRENCY` - Currency symbol

### **Adding More Users**
Add new rows to the **Employees** sheet:
1. Generate new Employee ID (EMP004, EMP005, etc.)
2. Enter Name, Role (Admin/Manager/Cashier), Username
3. For PIN: Use MD5 hash of desired PIN
4. For Password: Use SHA256 hash of desired password
5. Set Active = TRUE

### **Adding Products**
Add new rows to the **Products** sheet:
1. Generate new Product ID (PRD006, PRD007, etc.)
2. Enter Name, Price, Cost, Stock, Category
3. Generate or enter Barcode
4. Set Active = TRUE

---

## **📈 System Features Available**

### **✅ Authentication System**
- Dual PIN/Password login
- Virtual PIN pad for touch devices
- Role-based access control
- Session management

### **✅ Point of Sale**
- Product search and selection
- Shopping cart management
- Tax and discount calculations
- Multiple payment methods
- Receipt generation

### **✅ Inventory Management**
- Real-time stock tracking
- Product CRUD operations
- Category management
- Barcode support

### **✅ Advanced Reporting**
- **Daily RVC Report** - Revenue, Volume, Cost analysis
- **Weekly Sales Summary** - Trend analysis
- **Monthly Performance** - Comprehensive metrics
- **Product Performance** - Detailed product analytics
- **Employee Performance** - Staff productivity tracking

### **✅ Administrative Features**
- Employee management
- Settings configuration
- System logs
- Data export capabilities

---

## **🎯 Success Checklist**

- ✅ Google Spreadsheet created
- ✅ Apps Script project set up with all files
- ✅ `initializePOSSystem()` executed successfully
- ✅ `testSystemSetup()` shows success message
- ✅ Web app deployed and accessible
- ✅ Login working with default credentials
- ✅ All 5 sheets created with proper structure
- ✅ Sample data populated
- ✅ Reports system functional

---

## **📞 Support**

If you encounter any issues:

1. **Check the Apps Script console** for error messages
2. **Run `testSystemSetup()`** to verify system state
3. **Verify all file contents** were copied completely
4. **Check browser console** for frontend errors
5. **Try incognito/private browsing** to rule out cache issues

Your POS system is now ready for production use! 🎉