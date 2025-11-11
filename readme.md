# Excel Financial Data Connector (Index Constituent UDFs)

## Introduction

This project implements a robust and performant system that exposes powerful financial data retrieval capabilities directly within Microsoft Excel via User-Defined Functions (UDFs). The system is designed to allow financial analysts to query structured financial index constituent data (including weights, sector, and market capitalization category) using simple, formula-based syntax directly in an Excel cell, simulating a seamless integration experience.

#  Project Setup Guide

Follow these steps to set up and run the project locally.

---

### Clone the Repository

Clone the project from GitHub :

```bash
git clone https://github.com/Kiran3002/Excel_udf_itus.git
```

### install requirements.txt
```
$ pip install -r requirements.txt
```
### Edit config.ini file
#### Locate `config.ini` from project folder
Open the file and edit the database path:
> DB_PATH = add equity_index_constituents - nifty500.db database path from project folder
---

## install xlwings addin
xlwings lets you call Python functions directly from Excel, or manipulate Excel workbooks via Python code
**open cmd or bash**
installation of xlwings:
```
$ pip install xlwings
```
xlwings addin installation for Excel
```
$  xlwings addin install
```
Confirm xlwings Installation and Path
```
$ pip show xlwings
```
---
## Enable Trust Access to VBA Project Object Model

To allow `xlwings` to control Excel macros, you must enable access to the VBA project.

**Steps:**
1. Open Excel.
2. Go to:  
   `File → Options → Trust Center → Trust Center Settings → Macro Settings`
3. Check the box:  
   **"Trust access to the VBA project object model"**
4. Click **OK** to apply.

> This step is required for Python functions to interact with Excel macros using `xlwings`.
---

## Add the xlwings Excel Add-in Manually

If the `xlwings` tab is not visible in Excel, add it manually:

1. Open **Excel**.
2. Go to `File → Options → Add-ins`.
3. In the **Manage** dropdown (bottom), select **Excel Add-ins** and click **Go**.
4. Click **Browse** and navigate to your xlwings add-in file.
5. Select the **xlwings.xlam** file and click **OK**.
6. Ensure the **xlwings** checkbox is checked.

> This will add the xlwings tab to your Excel ribbon, allowing you to run Python code directly from Excel.
---

## Enable xlwings Reference in VBA Editor

To ensure Excel’s VBA environment recognizes the `xlwings` library:

1. Open **Excel**.
2. Press **`Alt + F11`** to open the **VBA Editor**.
3. Go to **Tools → References...**
4. Find **xlwings** in the list and check the box.
5. If it’s missing:
   - Click **Browse** and navigate to your xlwings add-in file.
   - Select the **xlwings.xlam** file and click **OK**.
6. Click **OK** to save and close.

> This step is required for VBA macros and Python scripts to work together using xlwings.
---

## Configure xlwings Settings in Excel

After enabling the xlwings add-in, configure these settings:
go to xlwings tab:
![xlwings ribbon](images/ribbon.jpg)
### 1. Set Python Interpreter
- Go to **xlwings tab → Interpreter**
- Verify the path to your Python executable:
- If missing add the python path.

### 2. Set Python Path
- Go to **xlwings tab → Python Path**
- Add your project folder path.

### 3. Set UDF Module
- Go to **xlwings tab → UDF Modules**
- Enter the Python file name (without `.py`) that contains your xlwings functions:
ebitda_margins_data_udf

>  These settings ensure Excel connects to the correct Python environment, project folder, and script for executing xlwings functions.
---

## click on import functions 

Once you’ve installed and configured `xlwings`, and clicked **Import Functions** from the xlwings tab in Excel,  
you can use the following custom formulas directly inside Excel — just like built-in Excel functions.

Each function pulls data from database and displays it as a formatted table.

---
## 🧩 Available Functions

#### 1. `get_monthly_data(index_name, date_value)`
Fetch constituents for a given index as on a specific date.
**Example Usage:**
```excel
=get_monthly_data("nifty_500", "2023-04-30")
```
output columns:
company_name | sector | mcap_category | weights

#### 2. `get_series(index_name, start_date, end_date)`
Fetch index constituents and their weights between two dates (inclusive).

**Example Usage:**
```excel
=get_series("nifty_50", "2020-03-31", "2025-09-30")
```
Output Columns:
index_name | accord_code | company_name | sector | mcap_category | date | weights
#### 3. `get_matrix(date_value, index_name)`
Fetch all constituents of a given index as on a specific date.

**Example Usage:**
```excel
=get_matrix("2023-04-30", "nifty_500")
```
output columns:
accord_code | company_name | sector | mcap_category | date | weights

#### 4. `get_all_data(index_name)`
Fetch all available data for a specific index across all dates.

**Example Usage:**
```excel
=get_all_data("nifty_500")
```
output columns:
accord_code | company_name | sector | mcap_category | date | weights

For any queries: raise an issue or feel free to reach out kirankumar300213@gmail.com
---




