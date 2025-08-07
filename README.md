# Modern POS System

A fully-featured, modern Point of Sale (POS) application built with Google Sheets, Google Apps Script, and HTML/CSS/JavaScript. This system provides a complete solution for inventory tracking, sales management, employee administration, and comprehensive reporting.

## 🚀 Features

### 🏪 Core POS Functionality
- **Product Search & Barcode Support**: Search products by name, category, or barcode
- **Shopping Cart**: Add/remove items with real-time calculations
- **Multiple Payment Methods**: Cash, Credit Card, Debit Card, Mobile Payment
- **Tax & Discount Calculations**: Configurable tax rates and discount rules
- **Stock Management**: Real-time inventory tracking with low-stock alerts
- **Receipt Generation**: Digital receipts with configurable footer text

### 👥 Employee Management
- **Role-Based Access Control**: Admin, Manager, and Cashier roles
- **Secure Authentication**: PIN-based login system with encrypted storage
- **Activity Logging**: Comprehensive audit trail of all system actions
- **Session Management**: Secure user sessions with automatic timeout

### 📊 Analytics & Reporting
- **Interactive Dashboard**: Real-time statistics and KPIs
- **Sales Analytics**: Daily, weekly, and monthly sales reports
- **Product Performance**: Top-selling products and category analysis
- **Payment Methods Distribution**: Visual breakdown of payment preferences
- **Key Metrics**: Revenue, items sold, average order value, and more

### 🎨 Modern UI/UX
- **Responsive Design**: Works on desktop, tablet, and mobile devices
- **CSS3 Animations**: Smooth transitions and hover effects
- **Material Design**: Clean, modern interface with FontAwesome icons
- **Keyboard Shortcuts**: Quick access to common functions
- **Real-time Updates**: Live data synchronization across all users

### ⚙️ Administration
- **Product Management**: Add, edit, and manage product catalog
- **Employee Administration**: User management with role assignment
- **System Settings**: Configurable tax rates, currency, and preferences
- **Data Export**: Export reports as PDF and CSV (coming soon)
- **System Logs**: Complete audit trail of all activities

## 🛠️ Technical Stack

- **Backend**: Google Apps Script (JavaScript)
- **Database**: Google Sheets
- **Frontend**: HTML5, CSS3, JavaScript (ES6+)
- **Charts**: Google Charts API
- **Icons**: FontAwesome 6.4.0
- **Fonts**: Inter (Google Fonts)
- **Authentication**: Custom PIN-based system with MD5 hashing

## 📋 Prerequisites

1. **Google Account**: Required for Google Sheets and Apps Script
2. **Google Sheets**: Create a new spreadsheet for the POS data
3. **Google Apps Script**: Access to script.google.com
4. **Modern Web Browser**: Chrome, Firefox, Safari, or Edge (latest versions)

## 🚀 Installation & Setup

### Step 1: Create Google Spreadsheet
1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Name it "Modern POS System"
4. Note the spreadsheet ID from the URL

### Step 2: Setup Apps Script Project
1. Go to [Google Apps Script](https://script.google.com)
2. Create a new project
3. Name it "Modern POS System"
4. Delete the default `Code.gs` content

### Step 3: Add Project Files
1. **Create `Code.gs`**: Copy the content from `Code.gs` in this repository
2. **Create `index.html`**: Copy the content from `index.html` in this repository
3. **Create `pos-script.html`**: Copy the content from `pos-script.html` in this repository
4. **Create `charts.html`**: Copy the content from `charts.html` in this repository
5. **Update `appsscript.json`**: Copy the content from `appsscript.json` in this repository

### Step 4: Configure Project
1. In Apps Script, go to **Project Settings**
2. Check "Show 'appsscript.json' manifest file in editor"
3. Update the manifest file with the provided configuration
4. Save all files

### Step 5: Initialize System
1. In the Apps Script editor, select the `initializePOSSystem` function
2. Click **Run** to execute the function
3. Grant necessary permissions when prompted
4. This will create all required sheets and sample data

### Step 6: Deploy Web App
1. Click **Deploy** > **New deployment**
2. Choose type: **Web app**
3. Execute as: **Me**
4. Who has access: **Anyone** (or restrict as needed)
5. Click **Deploy**
6. Copy the web app URL

### Step 7: Test System
1. Open the web app URL
2. Use default credentials:
   - Username: `admin`
   - PIN: `1234`
3. Explore all features and functionality

## 📖 User Guide

### 🔐 Login System
- **Admin**: Full access to all features
- **Manager**: Access to sales, reports, and product management
- **Cashier**: Access to POS interface only

### 💻 Point of Sale Interface
1. **Product Search**: Use the search bar to find products by name, category, or barcode
2. **Add to Cart**: Click on product cards to add items to the cart
3. **Modify Quantities**: Use +/- buttons to adjust item quantities
4. **Checkout**: Click checkout button and select payment method
5. **Process Sale**: Complete the transaction and generate receipt

### 📦 Product Management
1. Navigate to **Products** section
2. Click **Add Product** to create new items
3. Fill in product details (name, price, cost, stock, barcode)
4. Use the table to view and manage existing products

### 👥 Employee Management (Admin Only)
1. Navigate to **Employees** section
2. Click **Add Employee** to create new user accounts
3. Assign appropriate roles and set login credentials
4. Manage existing employees through the table interface

### 📊 Reports & Analytics
1. Navigate to **Reports** section
2. View interactive charts and analytics
3. Filter data by date ranges
4. Export reports for external use

### ⚙️ System Settings (Admin Only)
1. Navigate to **Settings** section
2. Configure tax rates, currency, and other preferences
3. Set low-stock thresholds and receipt footer text
4. Save changes to apply system-wide

## 🔧 Configuration

### Tax Settings
```javascript
// Default tax rate (8.75%)
updateSetting('TAX_RATE', '0.0875', 'Default tax rate');
```

### Low Stock Alerts
```javascript
// Alert when stock falls below 10 items
updateSetting('LOW_STOCK_THRESHOLD', '10', 'Low stock threshold');
```

### Receipt Customization
```javascript
// Custom footer for receipts
updateSetting('RECEIPT_FOOTER', 'Thank you for your business!', 'Receipt footer');
```

## 🛡️ Security Features

### Authentication
- PIN-based login system
- MD5 hash encryption for PINs
- Session management with automatic timeout
- Role-based access control

### Data Protection
- Server-side validation for all operations
- Audit logging for all user actions
- Secure communication between client and server
- Input sanitization and validation

### Access Control
- Different permission levels by role
- Hidden UI elements based on user permissions
- Server-side authorization checks
- Session-based security

## 📱 Mobile Support

The POS system is fully responsive and works on:
- **Desktop**: Full feature set with optimal layout
- **Tablet**: Touch-friendly interface with adapted layout
- **Mobile**: Compact view with essential features

### Mobile Features
- Touch-optimized buttons and controls
- Responsive grid layouts
- Collapsible navigation
- Optimized text sizes and spacing

## 🎯 Keyboard Shortcuts

- **Ctrl/Cmd + K**: Focus product search (in POS mode)
- **Escape**: Close open modals
- **Enter**: Submit forms
- **Tab**: Navigate between form fields

## 🔍 Troubleshooting

### Common Issues

**Login Not Working**
- Verify username and PIN are correct
- Check if user account is active
- Ensure proper permissions are granted

**Data Not Loading**
- Check internet connection
- Verify Google Sheets permissions
- Refresh the page and try again

**Charts Not Displaying**
- Ensure Google Charts library is loaded
- Check for JavaScript errors in browser console
- Verify sales data exists for the selected date range

**Performance Issues**
- Large datasets may cause slower loading
- Consider archiving old data to separate sheets
- Optimize by limiting date ranges in reports

### Error Messages

**"Authentication failed"**
- Check login credentials
- Verify user account exists and is active

**"Failed to load products"**
- Check Google Sheets permissions
- Verify sheet structure is correct

**"Error processing sale"**
- Check product stock availability
- Verify all required fields are completed

## 🚀 Advanced Features

### Barcode Integration
The system supports barcode input through:
1. **Barcode Scanner**: Connect USB/Bluetooth barcode scanner
2. **Manual Entry**: Type barcode numbers in search field
3. **Camera Scanning**: Integration ready for camera-based scanning

### Inventory Management
- **Automatic Stock Updates**: Stock decreases with each sale
- **Low Stock Alerts**: Notifications when items run low
- **Stock Adjustments**: Manual inventory corrections
- **Product Categories**: Organize products by type

### Reporting Features
- **Real-time Analytics**: Live data updates
- **Date Range Filtering**: Custom report periods
- **Export Capabilities**: PDF and CSV export (planned)
- **Visual Charts**: Interactive graphs and charts

## 🔮 Future Enhancements

### Planned Features
- [ ] **Camera Barcode Scanning**: Built-in camera support
- [ ] **Email Receipts**: Send receipts via email
- [ ] **Discount System**: Advanced discount rules and coupons
- [ ] **Multi-location Support**: Support for multiple store locations
- [ ] **Customer Management**: Customer accounts and loyalty programs
- [ ] **Integration APIs**: Connect with external systems
- [ ] **Offline Mode**: Continue operations without internet
- [ ] **Advanced Analytics**: Machine learning insights

### Enhancement Requests
- **Printer Integration**: Direct receipt printing
- **Payment Gateway**: Credit card processing
- **Backup & Restore**: Automated data backup
- **Multi-language**: Support for multiple languages
- **Theme Customization**: Custom color schemes and branding

## 🤝 Contributing

We welcome contributions to improve the POS system! Here's how you can help:

1. **Report Issues**: Submit bug reports and feature requests
2. **Code Contributions**: Fork the repository and submit pull requests
3. **Documentation**: Help improve documentation and guides
4. **Testing**: Test new features and provide feedback

### Development Guidelines
- Follow existing code style and conventions
- Add comments for complex logic
- Test thoroughly before submitting
- Update documentation for new features

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙋 Support

For support and questions:

1. **Documentation**: Check this README and code comments
2. **Issues**: Submit issues on the project repository
3. **Community**: Join discussions and share experiences
4. **Professional Support**: Contact for custom implementations

## 🎉 Acknowledgments

- **Google**: For Apps Script and Sheets platforms
- **FontAwesome**: For beautiful icons
- **Google Fonts**: For the Inter font family
- **Google Charts**: For interactive data visualization
- **Community**: For feedback and contributions

---

**Built with ❤️ for modern retail businesses**

*Transform your retail operations with this powerful, flexible POS system that grows with your business.*