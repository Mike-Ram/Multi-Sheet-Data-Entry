# Multi-Sheet-Data-Entry

A powerful, flexible Python-based data entry application with Excel integration and multi-sheet support. Perfect for automating repetitive data entry tasks across multiple categories or departments.

## Features
**Core Capabilities**
* **Multi-Sheet Support**: Create and manage up to 10 sheets in a single Excel file
* **Dynamic Column Configuration**: Define custom columns for each sheet via interactive CLI
* **Tabbed Interface**: Easy navigation between different data entry forms
* **Real-time Preview**: See your data as you enter it
* **Full Data Display**: View, update, and manage all records in dedicated windows
* **Flexible Setup**: Works with new or existing Excel files

## User Experience
* ✅ Intuitive GUI built with Tkinter
* ✅ Placeholder text in all input fields
* ✅ Smart submit button (enables only when all fields are filled)  
* ✅ Scrollable forms for handling many columns
* ✅ Error handling and validation
* ✅ Data refresh capabilities
* ✅ Clean, professional interface

## 📋 Table of Contents
* [Installation](#installation)
* [Quick Start](#quick-start)
* [Usage Guide](#usage-guide)
* [Configuration](#configuration)
* [Use Cases](#use-cases)
* [Screenshots](#screenshots)
* [Contributing](#contributing)
* [License](#license)

## 🔧 Installation
### Prerequisites
* Python 3.7 or higher
* pip (Python package installer)

### Required Libraries
```bash
pip install openpyxl
```
The following libraries are included with Python:
* **tkinter** (usually comes with Python)
* **pathlib**
* **os**
* **sys**

### Download
```bash
git clone https://github.com/yourusername/multi-sheet-data-entry.git
cd multi-sheet-data-entry
```

## 🚀 Quick Start
### First Time Setup
1. **Run the application:**
```bash
python data_entry.py
```

2. **Follow the interactive setup:**
   * Choose to create a new Excel file
   * Specify number of sheets (1-10)
   * Name each sheet
   * Define columns for each sheet
   * Select preview columns
   * Specify file path (or use default)

3. **Start entering data!**

### Using an Existing File

If you already have an Excel file:

1. Place your `.xlsx` file in the same directory or note its path
2. Run the application
3. Choose to use the existing file
4. Optionally edit the structure or start entering data
    
<img width="1366" height="768" alt="image1" src="https://github.com/user-attachments/assets/4f43b5b3-b920-4184-b4cc-2a68f1524464" />

**Or**
    
* Enter "no" in the command line.

<img width="1366" height="768" alt="image2" src="https://github.com/user-attachments/assets/722de078-b77a-4a9b-91de-64b6715fbbba" />


* Then, enter "yes". Select the file you've created for data entry.


<img width="1366" height="768" alt="image3" src="https://github.com/user-attachments/assets/09857cee-aa0b-4c20-af24-ab28a854d601" />


## 📖 Usage Guide

### Creating a New Excel File
```
============================================================
MULTI-SHEET DATA ENTRY SYSTEM - STARTUP
============================================================

============================================================
EXCEL FILE NOT FOUND - SETUP REQUIRED
============================================================

'sample.xlsx' does not exist.
Would you like to create it? (yes/no): yes
```

### Defining Sheet(s)
```
============================================================
SHEET SETUP
============================================================

How many sheets do you want to create? (1-10): 2
```
###### We entered 2 to create two sheets.

### Creating Sheet Names
```
Enter name for Sheet 1: Data

✓ Sheet name: Data
```
###### The name of sheet 1 is "Data".

### Defining Columns of Sheet(s)
```
--- Column Setup for 'Data' ---
Enter column names one by one. Press Enter with empty input to finish.
  Column 1: One Column
  ✓ Added: One Column
  Column 2: Two Column
  ✓ Added: Two Column
  Column 3: Three Column
  ✓ Added: Three Column
  Column 4:

Select columns to display in preview (Total available: 3):
  1. One Column
  2. Two Column
  3. Three Column

Enter column numbers separated by commas (e.g., 1,2,3,4)
Or press Enter to use first 4 columns.

Preview columns:
✓ Preview columns: One Column, Two Column, Three Column

✓ Sheet 'Data' configured successfully!
  - Columns: 3
  - Preview columns: 3
```
###### Follow the same process for sheet 2.

### Creating the Excel File
```
============================================================
CREATING EXCEL FILE
============================================================

Enter file path (press Enter for 'sample.xlsx'): C:\sample\sample.xlsx
```
###### Created the sample.xlsx file in the directory C:\sample\sample.xlsx

### Selecting Preview Columns
```
Select columns to display in preview (Total available: 4):
  1. Order ID
  2. Customer Name
  3. Product
  4. Quantity

Enter column numbers separated by commas (e.g., 1,2,3,4)
Or press Enter to use first 4 columns.

Preview columns: 1,2,3
✓ Preview columns: Order ID, Customer Name, Product
```
###### Preview: After you create the file, you can display selected columns of data on the right panel in the UI.

## Data Entry Interface

**The application opens with:**

* **Left Panel**: Data entry form with all fields
* **Right Panel**: Preview of entered data
* **Tabs**: One tab per sheet (if multiple sheets)
* **Buttons**: Submit, Clear, Full View

## Working with Data

### Adding Records:

1. Fill in all fields in the data entry form
2. Submit button becomes enabled when all fields are complete
3. Click "Submit" to save to Excel
4. Form clears automatically

### Viewing All Data:

1. Click "Full View" button
2. See all columns and records
3. Scroll horizontally/vertically as needed

### Updating Records:

1. Open "Full View"
2. Select a row
3. Click "Update Selected"
4. Modify fields in the popup window
5. Click "Save"

### Refreshing Data:

* Click "Refresh" in the Full View window to reload from Excel

## ⚙️ Configuration
### Default Settings
```python
DEFAULT_EXCEL_FILE = "Sample.xlsx"
```
You can change this in the code to use a different default filename.

### File Location
By default, the Excel file is created in the same directory as the script. You can specify a custom path during setup:
```
Enter file path (press Enter for 'Sample.xlsx'): C:/MyData/company_data.xlsx
```

## 💡 Use Cases
### 1. Automotive Service Center
```
Sheet 1: Pending Units
  - Unit No, Customer, VIN, Status, Technician...

Sheet 2: Completed Services
  - Service ID, Date, Total Cost, Parts Used...

Sheet 3: Parts Inventory
  - Part No, Description, Quantity, Supplier...
```

### 2. Restaurant Management
```
Sheet 1: Daily Orders
  - Order ID, Table, Items, Total, Server...

Sheet 2: Inventory
  - Ingredient, Stock Level, Reorder Point...

Sheet 3: Staff Schedule
  - Employee, Shift, Date, Hours...
```

### 3. School Administration
```
Sheet 1: Student Records
  - Student ID, Name, Grade, Section...

Sheet 2: Grades
  - Student ID, Subject, Score, Term...

Sheet 3: Attendance
  - Date, Student ID, Status, Remarks...
```

### 4. Sales Tracking
```
Sheet 1: Leads
  - Lead ID, Company, Contact, Status...

Sheet 2: Active Deals
  - Deal ID, Value, Stage, Close Date...

Sheet 3: Closed Sales
  - Sale ID, Revenue, Date, Salesperson...
```

## 📸 Screenshots

### Main Interface

<img width="1366" height="768" alt="image4" src="https://github.com/user-attachments/assets/2e584098-2e71-4fd3-9530-cf1a26bdd1e6" />


### Full Data Display Window

<img width="1366" height="768" alt="image5" src="https://github.com/user-attachments/assets/4bf28483-9e06-4f45-9a64-1e741a017da8" />

## 🛠️ Technical Details

### Classes
* **Config**: Manages Excel file configuration and sheet setup
* **PlaceholderEntry**: Custom Entry widget with placeholder text
* **UpdateWindow**: Window for editing existing records
* **DataDisplayWindow**: Full data display with update capabilities
* **SheetFrame**: Individual frame for each sheet's data entry
* **Window**: Main application window with tabs

### Key Technologies
* **GUI Framework**: Tkinter (ttk for modern widgets)
* **Excel Integration**: openpyxl
* **Data Display**: Treeview widgets with scrollbars


## 🐛 Troubleshooting
### Common Issues

**Issue**: "Excel file is open in another program"
* **Solution**: Close Excel before saving data in the application

**Issue**: Application won't start
* **Solution**: Ensure Python 3.7+ is installed and openpyxl is installed

**Issue**: Columns not displaying correctly
* **Solution**: Check that the first row of your Excel file contains column headers

**Issue**: Can't see all columns in preview
* **Solution**: Use "Full View" button to see all columns, or reconfigure preview columns

### Error Messages
* **"File not found"**: The Excel file was moved or deleted
* **"Permission denied"**: File is open in another application
* **"Invalid input"**: Check your column selections during setup

## 🔄 Updates and Versions
### Version 1.0.0 (Current)

* Initial release
* Multi-sheet support
* Dynamic column configuration
* Full CRUD operations
* Interactive CLI setup

## 🤝 Contributing
**Contributions are welcome! Here's how you can help:**

1. Fork the repository

2. Create a feature branch
```bash
git checkout -b feature/AmazingFeature
```

3. Commit your changes
```bash
git commit -m 'Add some AmazingFeature'
```

4. Push to the branch
```bash
git push origin feature/AmazingFeature
```

5. Open a Pull Request

### Contribution Guidelines

* Follow PEP 8 style guide for Python code
* Add comments for complex logic
* Update README.md if adding new features
* Test your changes thoroughly

## 📝 License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👤 Author

**Michael Ramalla**

* GitHub: [Mike-Ram](https://github.com/Mike-Ram)
* Email: MICHAEL.RAMALLA3011@gmail.com

## 🙏 Acknowledgments
* Built with Python and Tkinter
* Excel integration powered by openpyxl
* Inspired by the need to automate repetitive data entry tasks

## 📞 Support
**If you have any questions or need help:**

1. Check the [Troubleshooting](#troubleshooting) section
2. Open an [Issue](https://github.com/yourusername/multi-sheet-data-entry/issues)
3. Contact via email

## ⭐ Star This Repository
If you find this project useful, please consider giving it a star! It helps others discover the project.

