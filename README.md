```text
README.md
```

````markdown
# Buffet Report Automation

Python automation project for preparing and updating weekly Buffet Report data in Excel.

This project was created to reduce manual Excel work, including combining raw files, cleaning data, appending data into the main report workbook, updating formulas, refreshing pivot tables, filtering latest weeks, and saving the final report with the latest data date.

---

## Project Objective

The main objective of this project is to automate the weekly Buffet Report workflow and reduce repetitive manual tasks in Excel.

Before automation, the process required opening multiple files, running VBA manually, copying data into the main report file, updating formulas, refreshing pivot tables, and checking report sheets one by one.

This automation helps make the process faster, more consistent, and easier to control.

---

## Workflow Overview

1. Select the main Buffet Report Excel file
2. Back up the selected report file before processing
3. Read input folders from the Control sheet
4. Combine raw Excel files into working sheets
5. Run VBA cleaning logic
6. Append cleaned data into the main report workbook
7. Fill formulas down to the latest data row
8. Replace restaurant names for reporting consistency
9. Update week dropdown values in report tables
10. Refresh pivot tables
11. Filter database sheets to show the latest 5 weeks
12. Save As the report file using the latest data date

---

## Folder Structure

Example working folder:

```text
buffe/
│
├── Buffe_Final.py
├── Buffe_VBA.xlsm
├── 04_Buffet All Resort 260126-170526.xlsx
│
├── สำรอง/
│
├── SKC/
│   ├── ALL/
│   └── DINNER/
│
├── SNT/
│   ├── ALL/
│   └── DINNER/
│
├── SPN/
│   ├── ALL/
│   ├── DINNER/
│   └── LUNCH/
│
└── SYY/
    ├── ALL/
    └── DINNER/
````

---

## Control Sheet Example

The VBA workbook uses a Control sheet to define input folders by branch and data type.

| InputFolder                  | Branch | DataType |
| ---------------------------- | ------ | -------- |
| C:\Users...\buffe\SKC\ALL    | SKC    | ALL      |
| C:\Users...\buffe\SKC\DINNER | SKC    | DINNER   |
| C:\Users...\buffe\SNT\ALL    | SNT    | ALL      |
| C:\Users...\buffe\SNT\DINNER | SNT    | DINNER   |
| C:\Users...\buffe\SPN\ALL    | SPN    | ALL      |
| C:\Users...\buffe\SPN\DINNER | SPN    | DINNER   |
| C:\Users...\buffe\SPN\LUNCH  | SPN    | LUNCH    |
| C:\Users...\buffe\SYY\ALL    | SYY    | ALL      |
| C:\Users...\buffe\SYY\DINNER | SYY    | DINNER   |

---

## Main Features

### 1. File Backup

Before the automation starts, the selected main report file is copied into the backup folder.

This prevents data loss if an error occurs during processing.

### 2. Auto Combine Raw Files

The script loops through all input folders and combines Excel files into sheets for each branch and data type.

### 3. VBA Cleaning Process

The automation works together with existing VBA logic in Excel.

The Python script controls the workflow, while VBA handles the original cleaning rules.

### 4. Append Data to Main Report

Cleaned data is copied into the correct report sheets, such as:

* SKC All
* SKC Dinner
* SNT All
* SNT Dinner
* SPN All
* SPN Dinner
* SPN Lunch
* SYY All
* SYY Dinner

### 5. Formula Fill Down

After new data is added, formulas are filled down automatically based on the latest data row.

### 6. Restaurant Name Standardization

Some restaurant names are replaced for consistency, for example:

* All By The Sea Restaurant And Bar → By The Sea Restaurant And Bar
* All Chantara Restaurant → Chantara Restaurant
* All Saaitara Restaurant → Saaitara Restaurant

### 7. Week Dropdown Update

The script checks the latest week from the newly added data and adds it into dropdown source lists without deleting old weeks.

This allows report users to select the latest week from the table dropdown.

### 8. Pivot Refresh and Latest Week Filter

Pivot tables are refreshed, and database sheets are filtered to show the latest 5 weeks.

### 9. Auto Save As by Latest Date

The final report file is saved with the latest data date from the appended data.

Example:

```text
04_Buffet All Resort 260126-170526.xlsx
```

---

## Tools Used

* Python
* Excel VBA
* Microsoft Excel
* win32com / pywin32
* OpenPyXL
* Pandas

---

## Skills Demonstrated

* Excel automation
* VBA and Python integration
* Report workflow automation
* Data cleaning
* Data validation
* Pivot table refresh automation
* Folder-based data processing
* Error handling and backup control
* Hospitality reporting process improvement

---

## Result

This project helps reduce manual reporting work and improves report consistency.

The workflow can be run from Python, while still keeping the existing Excel and VBA process that users are familiar with.

---

## Notes

This repository uses mock or sanitized data only.

Real business data, supplier information, financial details, and internal company files are not included.

```

