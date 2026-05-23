```markdown
# Buffet Report Automation

This project automates the weekly buffet report workflow using Python and Excel VBA.

The original process required manually combining multiple Excel files by resort and meal type, cleaning the data, copying it into the main report file, refreshing pivot tables, updating weekly dropdown filters, and saving the report with the latest data date.

This automation reduces manual work, prevents date and week selection errors, and standardizes the report preparation process.

---

## Project Objective

To automate the weekly buffet reporting workflow for multiple resorts and data types.

The automation helps with:

- Combining source Excel files from multiple folders
- Cleaning and preparing data through existing VBA logic
- Appending cleaned data into the main report workbook
- Pulling formulas down to new rows
- Standardizing restaurant names
- Updating week dropdowns in report tables
- Refreshing pivot tables
- Filtering the latest 5 weeks
- Saving the report using the latest data date
- Creating a backup before processing

---

## Workflow Overview

1. Select the main report file before running the process.
2. Backup the selected main report file.
3. Check and reset the VBA workbook if needed.
4. Read the folder paths from the Control sheet.
5. Combine all Excel files from each folder into separate sheets.
6. Run the existing VBA cleaning workflow.
7. Copy cleaned data into the matching sheet in the main report file.
8. Fill down formulas for each report sheet.
9. Replace restaurant names for reporting consistency.
10. Add process tracking information.
11. Update weekly dropdown values.
12. Refresh pivot tables.
13. Filter the latest 5 weeks.
14. Save As the final report using the latest data date.

---

## Folder Structure Example

```text
buffe/
│
├── Buffe_Final.py
├── Buffe_VBA.xlsm
├── backup/
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

## Control Sheet Setup

The VBA workbook contains a `Control` sheet that stores the folder paths and mapping information.

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

## Tools Used

* Python
* pywin32 / win32com
* Microsoft Excel
* Excel VBA
* Pivot Table
* Excel formulas
* Data validation dropdowns

---

## Key Features

### 1. Automated File Combination

The script reads all Excel files inside each input folder and combines them into the VBA workbook as separate sheets.

### 2. Existing VBA Logic Integration

The automation keeps the original VBA cleaning logic and runs it through Python instead of manually clicking each macro.

### 3. Main Report Update

Cleaned data is copied into the correct sheet in the main report workbook based on branch and data type.

### 4. Formula Fill Down

The script automatically fills formulas down to the latest appended rows.

Formula columns differ by sheet:

| Sheet Type  | Formula Columns |
| ----------- | --------------- |
| ALL / LUNCH | J:N             |
| DINNER      | J:Q             |

### 5. Restaurant Name Standardization

Some restaurant names are standardized after data is inserted, for example:

| Original Name                     | Standardized Name             |
| --------------------------------- | ----------------------------- |
| All By The Sea Restaurant And Bar | By The Sea Restaurant And Bar |
| All Chantara Restaurant           | Chantara Restaurant           |
| All Saaitara Restaurant           | Saaitara Restaurant           |

### 6. Week Dropdown Update

The script checks the latest week from the inserted data and adds it into dropdown sources without deleting old week values.

### 7. Pivot Refresh and Latest Week Filter

Pivot tables are refreshed and filtered to show only the latest 5 weeks.

### 8. Auto Save As

The final report is saved with the latest data date from the inserted data.

Example:

```text
04_Buffet All Resort 260126-170526.xlsx
```

---

## Business Impact

Before automation, the report preparation required many manual Excel steps and was at risk of human error, especially when copying data, updating formulas, selecting weeks, and saving the file name.

After automation, the workflow became more consistent and easier to repeat every report cycle.

Main improvements:

* Reduced manual copy-paste work
* Reduced risk of wrong week selection
* Reduced risk of incorrect report file name
* Improved report consistency
* Created backup before editing files
* Made the process easier to hand over and reuse

---

## Notes

This project was designed for an internal weekly buffet report workflow.

Mock data is used for portfolio demonstration. Real company data is not included in this repository.

````

```text
Add README for buffet report automation project
````

