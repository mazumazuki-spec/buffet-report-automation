# Buffe.py - Final automation version
# Updated for: append, formula fill, restaurant replace log, dropdown week, pivot latest 5 weeks, save as latest date

import os
import re
import shutil
from datetime import datetime, timedelta
import win32com.client as win32

# =====================================================
# 1) PATH SETTING
# =====================================================

BASE_FOLDER = r"C:\Users\STYNB113\Desktop\Python\buffe"

VBA_FILE = os.path.join(BASE_FOLDER, "Buffe_VBA.xlsm")

# MAIN_FILE จะให้เลือกเองตอนเริ่มรัน
# เพื่อตัดปัญหาชื่อไฟล์ Report เปลี่ยนทุกวีค
MAIN_FILE = None

# อ่าน Week จากคอลัมน์นี้ หลังจากลากสูตรแล้ว
WEEK_SOURCE_COLUMN = "L"

# อ่านวันที่จากคอลัมน์นี้ของข้อมูลที่เพิ่ง Append เข้า Report
DATE_SOURCE_COLUMN = "I"

# =====================================================
# 2) VBA MACRO NAMES
# =====================================================

MACRO_RESET = "ResetSheets_Safe_Auto"
MACRO_COMBINE = "CombineFilesToSheets_Auto"
MACRO_CLEAN = "Run_Cleaning_Buffe_Auto"

# Excel constants
XL_UP = -4162
XL_TO_LEFT = -4159
XL_CELL_TYPE_VISIBLE = 12
XL_PAGE_FIELD = 3


def normalize_excel_path(path):
    """
    Normalize path ให้ Excel COM ใช้งานได้เสถียรกว่า
    กันเคส path ปน / กับ \\ แล้ว Excel ไปอ่านเป็น C:\\//Users/...
    """

    if path is None:
        return None

    path = os.path.abspath(os.path.normpath(str(path)))
    return path.replace("/", "\\")


def get_excel_file_format(file_path):
    """
    FileFormat สำหรับ SaveAs ผ่าน Excel COM
    .xlsx = 51, .xlsm = 52
    """

    ext = os.path.splitext(str(file_path))[1].lower()

    if ext == ".xlsm":
        return 52

    return 51


# =====================================================
# 3) LOG FUNCTION
# =====================================================

def log(message):
    now = datetime.now().strftime("%H:%M:%S")
    print(f"[{now}] {message}")


def select_main_file():
    """
    ให้เลือกไฟล์ Report หลักก่อนเริ่มรัน
    ใช้ไฟล์ที่เลือกนี้ต่อทั้ง Backup / Append / Refresh / Save As
    """

    import tkinter as tk
    from tkinter import filedialog

    log("Waiting for user to select Main Report file...")

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    file_path = filedialog.askopenfilename(
        title="เลือกไฟล์ Report หลัก เช่น 04_Buffet All Resort 260126-170526.xlsx",
        initialdir=BASE_FOLDER,
        filetypes=[
            ("Excel files", "*.xlsx *.xlsm"),
            ("All files", "*.*"),
        ],
    )

    root.destroy()

    if not file_path:
        raise Exception("ไม่ได้เลือกไฟล์ Report หลัก")

    file_path = normalize_excel_path(file_path)

    log(f"✓ Selected Main File: {file_path}")

    return file_path


def backup_files_before_start():
    """
    สำรองไฟล์ก่อนเริ่มทำงานทุกอย่าง
    - สร้างโฟลเดอร์ BASE_FOLDER\สำรอง ถ้ายังไม่มี
    - สำรองไฟล์หลัก Report ที่กำลังจะถูกแก้
    - สำรองไฟล์ Buffe_VBA ที่มีการเขียน Status / Reset / Control
    - ใส่ timestamp กันชื่อซ้ำ ไม่ทับไฟล์สำรองเดิม
    """

    backup_folder = os.path.join(BASE_FOLDER, "สำรอง")
    os.makedirs(backup_folder, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    files_to_backup = [
        MAIN_FILE,
        VBA_FILE,
    ]

    log("Backing up files before start...")
    log(f"Backup folder: {backup_folder}")

    for file_path in files_to_backup:
        if not file_path:
            log("⚠ ไม่พบ Path สำหรับ Backup")
            continue

        if not os.path.isfile(file_path):
            log(f"⚠ ไม่พบไฟล์สำหรับ Backup: {file_path}")
            continue

        folder, file_name = os.path.split(file_path)
        name, ext = os.path.splitext(file_name)
        backup_file = os.path.join(backup_folder, f"{name}_backup_{timestamp}{ext}")

        shutil.copy2(file_path, backup_file)
        log(f"✓ Backup completed: {backup_file}")


# =====================================================
# 4) BASIC FUNCTIONS
# =====================================================

def run_macro(excel, macro_name):
    macro_full_name = f"'{os.path.basename(VBA_FILE)}'!{macro_name}"
    log(f"▶ Running Macro: {macro_name}")
    excel.Application.Run(macro_full_name)
    log(f"✓ Finished Macro: {macro_name}")


def reset_if_needed(excel, wb):
    log("Checking remaining sheets before start...")

    sheet_names = [ws.Name for ws in wb.Worksheets]
    other_sheets = [s for s in sheet_names if s != "Control"]

    log(f"Current Sheets: {sheet_names}")

    if other_sheets:
        log(f"พบชีทค้าง: {other_sheets}")
        run_macro(excel, MACRO_RESET)
        wb.Save()
        log("✓ Reset เรียบร้อย")
    else:
        log("✓ ไม่มีชีทค้าง มีแค่ Control")


def force_reset(excel, wb):
    log("Reset workbook after current round...")
    run_macro(excel, MACRO_RESET)
    wb.Save()
    log("✓ Reset หลังจบรอบเรียบร้อย")


def get_control_rows(ws_control):
    log("Reading Control rows...")

    rows = []
    last_row = ws_control.Cells(ws_control.Rows.Count, 1).End(XL_UP).Row

    log(f"Control last row: {last_row}")

    for r in range(2, last_row + 1):
        input_folder = ws_control.Cells(r, 1).Value
        branch = ws_control.Cells(r, 2).Value
        data_type = ws_control.Cells(r, 3).Value

        if input_folder and branch and data_type:
            rows.append({
                "row": r,
                "input_folder": str(input_folder).strip(),
                "branch": str(branch).strip(),
                "data_type": str(data_type).strip()
            })

            log(
                f"Found Control Row {r}: "
                f"{str(branch).strip()} | "
                f"{str(data_type).strip()} | "
                f"{str(input_folder).strip()}"
            )

    return rows


def set_current_area(ws_control, input_folder, branch, data_type):
    log("Setting Current Area in Control I1:I3...")

    ws_control.Range("I1").Value = input_folder
    ws_control.Range("I2").Value = branch
    ws_control.Range("I3").Value = data_type

    log(f"I1 CurrentInputFolder = {input_folder}")
    log(f"I2 CurrentBranch      = {branch}")
    log(f"I3 CurrentDataType    = {data_type}")


def update_status(ws_control, row, status, remark=""):
    ws_control.Cells(row, 4).Value = status
    ws_control.Cells(row, 5).Value = remark

    log(f"Update Status Row {row}: {status} | {remark}")


def count_xlsx_files(input_folder):
    if not os.path.isdir(input_folder):
        return 0

    files = [
        f for f in os.listdir(input_folder)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    ]

    log(f"Files found in folder: {len(files)}")

    for f in files:
        log(f"  - {f}")

    return len(files)


# =====================================================
# 5) AFTER APPEND: FILL FORMULA
# =====================================================

def fill_formulas_after_append(ws_target, paste_row, row_count, target_sheet_name):
    """
    ลากสูตรลงหลัง Append ข้อมูล
    ใช้สูตรจากแถวก่อนหน้า แล้ว AutoFill ลงมาถึงแถวสุดท้ายที่ Append
    """

    formula_map = {
        "SPN Lunch": ("J", "N"),
        "SPN Dinner": ("J", "Q"),
        "SPN All": ("J", "N"),

        "SYY All": ("J", "N"),
        "SYY Dinner": ("J", "Q"),

        "SNT Dinner": ("J", "Q"),
        "SNT All": ("J", "N"),

        "SKC All": ("J", "N"),
        "SKC Dinner": ("J", "Q"),
    }

    log(f"Checking formula fill config for: {target_sheet_name}")

    if target_sheet_name not in formula_map:
        log(f"⚠ ไม่มี Config สูตรสำหรับชีท: {target_sheet_name}")
        return

    start_col_letter, end_col_letter = formula_map[target_sheet_name]

    start_col = ws_target.Range(f"{start_col_letter}1").Column
    end_col = ws_target.Range(f"{end_col_letter}1").Column

    first_new_row = paste_row
    last_new_row = paste_row + row_count - 1
    formula_source_row = paste_row - 1

    log(
        f"Formula range config: {target_sheet_name} | "
        f"{start_col_letter}:{end_col_letter}"
    )
    log(f"Formula source row: {formula_source_row}")
    log(f"Formula fill rows: {first_new_row} to {last_new_row}")

    if formula_source_row < 2:
        log(f"⚠ ข้ามการลากสูตร: ไม่มีแถวสูตรก่อนหน้าใน {target_sheet_name}")
        return

    source_range = ws_target.Range(
        ws_target.Cells(formula_source_row, start_col),
        ws_target.Cells(formula_source_row, end_col)
    )

    fill_range = ws_target.Range(
        ws_target.Cells(formula_source_row, start_col),
        ws_target.Cells(last_new_row, end_col)
    )

    source_range.AutoFill(Destination=fill_range)

    log(
        f"✓ Fill Formula สำเร็จ: {target_sheet_name} | "
        f"{start_col_letter}:{end_col_letter} | "
        f"Rows {first_new_row}-{last_new_row}"
    )


# =====================================================
# 6) READ WEEK AFTER FORMULA FILL
# =====================================================

def get_week_from_appended_rows(ws_target, paste_row, row_count, week_col_letter="L"):
    """
    อ่านค่า Week จากข้อมูลที่เพิ่ง Append เข้าไป
    ใช้หลังจากลากสูตรแล้ว
    """

    week_col = ws_target.Range(f"{week_col_letter}1").Column

    first_row = paste_row
    last_row = paste_row + row_count - 1

    log(
        f"Reading Week from {ws_target.Name}!"
        f"{week_col_letter}{first_row}:{week_col_letter}{last_row}"
    )

    week_range = ws_target.Range(
        ws_target.Cells(first_row, week_col),
        ws_target.Cells(last_row, week_col)
    )

    values = week_range.Value

    if values is None:
        log("⚠ ไม่พบค่า Week ในช่วงข้อมูลที่ Append")
        return None

    if not isinstance(values[0], tuple):
        values = (values,)

    found_weeks = []

    for row in values:
        week_value = row[0]

        if week_value is not None and str(week_value).strip() != "":
            found_weeks.append(str(week_value).strip())

    if not found_weeks:
        log("⚠ Week column ว่างทั้งหมด")
        return None

    report_week = found_weeks[-1]

    log(f"✓ Found Report Week: {report_week}")

    return report_week


# =====================================================
# 7) AFTER APPEND: CLEAN RESTAURANT NAME + LOG
# =====================================================

def clean_restaurant_names(ws_target, target_sheet_name, paste_row, row_count):
    """
    Clean ชื่อ Restaurant ในคอลัมน์ H เฉพาะแถวที่เพิ่ง Append เข้ามา
    และบันทึก Log ในคอลัมน์ S ว่า Replace จากอะไรเป็นอะไร
    """

    restaurant_rules = {
        "SPN Lunch": {
            "All By The Sea Restaurant And Bar": "By The Sea Restaurant And Bar",
        },

        "SPN Dinner": {
            "All By The Sea Restaurant And Bar": "By The Sea Restaurant And Bar",
            "All Chantara Restaurant": "Chantara Restaurant",
        },

        "SPN All": {
            "All By The Sea Restaurant And Bar": "By The Sea Restaurant And Bar",
            "All Chantara Restaurant": "Chantara Restaurant",
        },

        "SYY All": {
            "All By The Sea Restaurant And Bar": "By The Sea Restaurant And Bar",
            "All Chantara Restaurant": "Chantara Restaurant",
            "All Saaitara Restaurant": "Saaitara Restaurant",
        },

        "SYY Dinner": {
            "All By The Sea Restaurant And Bar": "By The Sea Restaurant And Bar",
            "All Chantara Restaurant": "Chantara Restaurant",
            "All Saaitara Restaurant": "Saaitara Restaurant",
        },

        "SNT Dinner": {
            "All By The Sea Restaurant And Bar": "By The Sea Restaurant And Bar",
            "All By The Sea": "By The Sea",
        },

        "SNT All": {
            "All By The Sea Restaurant And Bar": "By The Sea Restaurant And Bar",
            "All By The Sea": "By The Sea",
        },

        "SKC All": {
            "All By The Sea Restaurant And Bar": "By The Sea Restaurant And Bar",
            "All Chantara Restaurant": "Chantara Restaurant",
        },

        "SKC Dinner": {
            "All By The Sea Restaurant And Bar": "By The Sea Restaurant And Bar",
            "All Chantara Restaurant": "Chantara Restaurant",
        },
    }

    log(f"Checking restaurant clean rules for: {target_sheet_name}")

    ws_target.Cells(1, 19).Value = "Restaurant Replace Log"

    if target_sheet_name not in restaurant_rules:
        log(f"⚠ ไม่มี Rule ร้านอาหารสำหรับชีท: {target_sheet_name}")
        return

    rules = restaurant_rules[target_sheet_name]

    first_row = paste_row
    last_row = paste_row + row_count - 1

    log(f"Restaurant clean range: H{first_row}:H{last_row}")
    log(f"Restaurant log range: S{first_row}:S{last_row}")

    restaurant_range = ws_target.Range(
        ws_target.Cells(first_row, 8),
        ws_target.Cells(last_row, 8)
    )

    values = restaurant_range.Value

    if values is None:
        log("⚠ Restaurant column is empty")
        return

    if not isinstance(values[0], tuple):
        values = (values,)

    new_values = []
    log_values = []
    replace_count = 0

    for row in values:
        old_value_raw = row[0]

        if old_value_raw is None:
            old_value = ""
        else:
            old_value = str(old_value_raw).strip()

        if old_value in rules:
            new_value = rules[old_value]
            replace_text = f"{old_value} -> {new_value}"
            replace_count += 1
        else:
            new_value = old_value_raw
            replace_text = ""

        new_values.append((new_value,))
        log_values.append((replace_text,))

    restaurant_range.Value = tuple(new_values)

    ws_target.Range(
        ws_target.Cells(first_row, 19),
        ws_target.Cells(last_row, 19)
    ).Value = tuple(log_values)

    log(
        f"✓ Clean Restaurant สำเร็จ: {target_sheet_name} | "
        f"Replace {replace_count} cells"
    )


# =====================================================
# 8) UPDATE TABLE WEEK DROPDOWN
# =====================================================

def _quote_sheet_name(sheet_name):
    return "'" + str(sheet_name).replace("'", "''") + "'"


def _range_address_external(rng):
    return f"={_quote_sheet_name(rng.Worksheet.Name)}!{rng.Address}"


def _flatten_first_col(values):
    result = []

    if values is None:
        return result

    if isinstance(values, tuple):
        for row in values:
            if isinstance(row, tuple):
                result.append("" if row[0] is None else str(row[0]).strip())
            else:
                result.append("" if row is None else str(row).strip())
    else:
        result.append(str(values).strip())

    return result


def _resolve_dropdown_source_range(wb, ws, formula1_text):
    """
    หา Source Range จริงของ Dropdown ให้ได้มากที่สุด
    รองรับ:
    - =WeekList / Named Range ระดับ Workbook
    - =Table SPN!$A$1:$A$20 หรือ ='Table SPN'!$A$1:$A$20
    - Named Range ระดับ Sheet
    - Formula ที่ Excel Evaluate แล้วได้ Range
    """

    ref_text = str(formula1_text).strip()
    if ref_text.startswith("="):
        ref_text = ref_text[1:].strip()

    app = wb.Application

    # 1) ลองให้ Excel resolve โดยตรงก่อน
    try:
        rng = app.Range(ref_text)
        return rng, ref_text, "range"
    except Exception:
        pass

    # 2) Workbook Named Range
    try:
        nm = wb.Names(ref_text)
        rng = nm.RefersToRange
        return rng, ref_text, "workbook_name"
    except Exception:
        pass

    # 3) Sheet Named Range
    try:
        nm = ws.Names(ref_text)
        rng = nm.RefersToRange
        return rng, ref_text, "sheet_name"
    except Exception:
        pass

    # 4) Evaluate จาก Sheet ก่อน แล้วค่อยจาก Application
    try:
        rng = ws.Evaluate("=" + ref_text)
        if rng is not None:
            return rng, ref_text, "evaluated"
    except Exception:
        pass

    try:
        rng = app.Evaluate("=" + ref_text)
        if rng is not None:
            return rng, ref_text, "evaluated"
    except Exception:
        pass

    raise Exception(f"Resolve Dropdown Source ไม่ได้: {formula1_text}")


def _set_validation_formula(cell, formula1):
    """ตั้งค่า Data Validation ใหม่แบบ List โดยยังไม่ยุ่งกับ Logic อื่น"""
    validation = cell.Validation
    validation.Delete()
    cell.Validation.Add(
        Type=3,
        AlertStyle=1,
        Operator=1,
        Formula1=formula1
    )


def ensure_dropdown_has_value(ws, cell_address, value):
    """
    เช็ก Dropdown ของ Cell ว่ามี value ไหม
    ถ้าไม่มี ให้เพิ่ม value เข้า Source จริงของ Dropdown โดยไม่ลบ Week เก่า

    เวอร์ชันแก้ล่าสุด:
    - หาว่า Dropdown ตั้ง Source มาจากไหนจริง ๆ ก่อน
    - ถ้า Source เป็น Range / Named Range: เพิ่ม Week ลงใน Range นั้น
    - ถ้า Source เต็ม: เพิ่มต่อท้าย แล้วขยาย Range / Named Range ให้ครอบคลุม Week ใหม่
    - ไม่สร้าง List ใหม่จากค่าที่อ่านได้ เพราะจะเสี่ยงทำ Week เก่าหาย
    - สุดท้าย Set Cell เป็น Week ล่าสุดด้วย
    """

    cell = ws.Range(cell_address)
    value = str(value).strip()

    if value == "":
        log(f"⚠ Week value ว่าง ข้าม Dropdown update: {ws.Name}!{cell_address}")
        return

    try:
        validation = cell.Validation
        formula1 = validation.Formula1
    except Exception:
        log(f"⚠ {ws.Name}!{cell_address} ไม่มี Data Validation หรืออ่านไม่ได้")
        cell.Value = value
        log(f"✓ ใส่ค่า {value} ที่ {ws.Name}!{cell_address} โดยตรง")
        return

    if formula1 is None or str(formula1).strip() == "":
        log(f"⚠ {ws.Name}!{cell_address} Dropdown ไม่มี Formula1")
        cell.Value = value
        log(f"✓ ใส่ค่า {value} ที่ {ws.Name}!{cell_address} โดยตรง")
        return

    formula1_text = str(formula1).strip()
    log(f"Dropdown Formula at {ws.Name}!{cell_address}: {formula1_text}")

    # -----------------------------
    # Case 1: Dropdown เป็น List ตรง ๆ เช่น Week 12,Week 13
    # -----------------------------
    if not formula1_text.startswith("="):
        items = [x.strip() for x in formula1_text.split(",") if x.strip()]

        if value not in items:
            items.append(value)
            new_formula = ",".join(items)
            _set_validation_formula(cell, new_formula)
            log(f"✓ เพิ่ม {value} เข้า Dropdown List เดิม: {ws.Name}!{cell_address}")
        else:
            log(f"✓ Dropdown List มี {value} อยู่แล้ว: {ws.Name}!{cell_address}")

        cell.Value = value
        log(f"✓ Set {ws.Name}!{cell_address} = {value}")
        return

    # -----------------------------
    # Case 2: Dropdown อ้างอิง Range / Named Range
    # -----------------------------
    try:
        wb = ws.Parent
        ref_range, ref_name, source_type = _resolve_dropdown_source_range(wb, ws, formula1_text)

        log(
            f"✓ Dropdown Source จริง: {ref_range.Worksheet.Name}!{ref_range.Address} "
            f"| Type: {source_type}"
        )

        existing_values = _flatten_first_col(ref_range.Value)
        existing_values_no_blank = [x for x in existing_values if x != ""]

        if value in existing_values_no_blank:
            log(f"✓ Dropdown Source มี {value} อยู่แล้ว")
            cell.Value = value
            log(f"✓ Set {ws.Name}!{cell_address} = {value}")
            return

        # 1) ถ้ามีช่องว่างใน Source เดิม ให้ใส่ช่องว่างนั้นก่อน ไม่ต้องขยาย Range
        blank_offset = None
        for i, existing_value in enumerate(existing_values, start=1):
            if existing_value == "":
                blank_offset = i
                break

        if blank_offset is not None:
            target_cell = ref_range.Cells(blank_offset, 1)
            target_cell.Value = value
            log(
                f"✓ เพิ่ม {value} ในช่องว่างของ Dropdown Source: "
                f"{target_cell.Worksheet.Name}!{target_cell.Address}"
            )

            # Source เดิมครอบอยู่แล้ว ไม่ต้องแก้ Validation / Named Range
            cell.Value = value
            log(f"✓ Set {ws.Name}!{cell_address} = {value}")
            return

        # 2) ถ้า Source เต็ม ให้เพิ่มต่อท้าย แล้วขยาย Source ให้คุมถึง Cell ใหม่
        target_cell = ref_range.Cells(ref_range.Rows.Count, 1).Offset(1, 0)
        target_cell.Value = value

        first_cell = ref_range.Cells(1, 1)
        expanded_range = ref_range.Worksheet.Range(first_cell, target_cell)

        log(
            f"✓ เพิ่ม {value} ต่อท้าย Dropdown Source: "
            f"{target_cell.Worksheet.Name}!{target_cell.Address}"
        )

        new_formula = _range_address_external(expanded_range)

        if source_type == "workbook_name":
            try:
                wb.Names(ref_name).RefersTo = new_formula
                log(f"✓ ขยาย Workbook Named Range {ref_name} เป็น {new_formula}")
            except Exception as e:
                log(f"⚠ ขยาย Workbook Named Range ไม่สำเร็จ: {ref_name} | {e}")
                _set_validation_formula(cell, new_formula)
                log(f"✓ เปลี่ยน Dropdown Source เป็น {new_formula}")

        elif source_type == "sheet_name":
            try:
                ws.Names(ref_name).RefersTo = new_formula
                log(f"✓ ขยาย Sheet Named Range {ref_name} เป็น {new_formula}")
            except Exception as e:
                log(f"⚠ ขยาย Sheet Named Range ไม่สำเร็จ: {ref_name} | {e}")
                _set_validation_formula(cell, new_formula)
                log(f"✓ เปลี่ยน Dropdown Source เป็น {new_formula}")

        else:
            # Range ปกติ / Evaluate: ต้องแก้ Validation ของ Cell นี้ให้ Source ขยาย
            _set_validation_formula(cell, new_formula)
            log(f"✓ ขยาย Dropdown Source เป็น {new_formula}")

        cell.Value = value
        log(f"✓ Set {ws.Name}!{cell_address} = {value}")

    except Exception as e:
        log(f"⚠ เพิ่มค่าเข้า Dropdown Source ไม่สำเร็จที่ {ws.Name}!{cell_address}: {e}")
        cell.Value = value
        log(f"✓ ใส่ค่า {value} ที่ {ws.Name}!{cell_address} โดยตรงแทน")

def update_table_week_dropdowns(excel, main_file, week_value):
    """
    อัปเดต Week Dropdown ในชีท Table ตาม Cell ที่กำหนด
    """

    log("Opening Main File for Table Week Dropdown update...")
    main_wb = excel.Workbooks.Open(normalize_excel_path(main_file))

    try:
        week_cell_map = {
            "Table SNT": [
                "E149",
                "E315",
            ],

            "Table SKC": [
                "E17",
                "E33",
                "E336",
                "E354",
            ],

            "Table SPN": [
                "E69",
                "E299",
                "E328",
                "E559",
                "E587",
            ],

            "Table SYY": [
                "E267",
                "E293",
                "E310",
                "E434",
                "E458",
                "E475",
            ],
        }

        for sheet_name, cell_list in week_cell_map.items():
            log("=" * 80)
            log(f"Updating Week Dropdown Sheet: {sheet_name}")

            try:
                ws = main_wb.Worksheets(sheet_name)
            except Exception:
                log(f"⚠ ไม่เจอชีท: {sheet_name}")
                continue

            for cell_address in cell_list:
                log(f"Updating {sheet_name}!{cell_address} to {week_value}")
                ensure_dropdown_has_value(ws, cell_address, week_value)

        log("Saving Main File after Table Week Dropdown update...")
        main_wb.Save()
        log("✓ Table Week Dropdown update completed")

    finally:
        log("Closing Main File after Table Week Dropdown update...")
        main_wb.Close(SaveChanges=True)
        log("✓ Main File closed")


# =====================================================
# 9) COPY SUMMARY TO MAIN FILE
# =====================================================

def append_summary_to_main_file(excel, vba_wb, main_file, branch, data_type):
    """
    Copy data from Buffe_VBA Summary to Main Buffet file.
    - Skip Header
    - Paste to target sheet เช่น SKC All
    - Stamp run date in Column R
    - Header Column R = Py Date
    - Header Column S = Restaurant Replace Log
    - Fill formula after append
    - Read Week from Column L
    - Clean restaurant name in Column H
    """

    log("Opening Main File...")
    main_wb = excel.Workbooks.Open(normalize_excel_path(main_file))

    try:
        ws_summary = vba_wb.Worksheets("Summary")

        target_sheet_name = f"{branch} {data_type.title()}"

        log(f"Looking for target sheet: {target_sheet_name}")

        ws_target = None

        for ws in main_wb.Worksheets:
            if ws.Name.strip().lower() == target_sheet_name.strip().lower():
                ws_target = ws
                break

        if ws_target is None:
            raise Exception(f"ไม่เจอชีทปลายทางในไฟล์หลัก: {target_sheet_name}")

        log(f"✓ Target Sheet Found: {ws_target.Name}")

        # -----------------------------
        # Source Range: Summary A2:last
        # -----------------------------
        source_last_row = ws_summary.Cells(ws_summary.Rows.Count, 1).End(XL_UP).Row
        source_last_col = ws_summary.Cells(1, ws_summary.Columns.Count).End(XL_TO_LEFT).Column

        log(f"Summary last row: {source_last_row}")
        log(f"Summary last col: {source_last_col}")

        if source_last_row < 2:
            raise Exception("Summary ไม่มีข้อมูลให้ Copy")

        source_range = ws_summary.Range(
            ws_summary.Cells(2, 1),
            ws_summary.Cells(source_last_row, source_last_col)
        )

        data = source_range.Value

        if data is None:
            raise Exception("ไม่พบข้อมูลใน Summary")

        if not isinstance(data[0], tuple):
            data = (data,)

        row_count = len(data)
        col_count = len(data[0])

        log(f"Rows to append from Summary: {row_count}")
        log(f"Columns to append from Summary: {col_count}")

        # -----------------------------
        # Find Paste Row by Column A
        # -----------------------------
        target_last_row = ws_target.Cells(ws_target.Rows.Count, 1).End(XL_UP).Row

        if ws_target.Range("A1").Value in [None, ""]:
            paste_row = 1
        else:
            paste_row = target_last_row + 1

        log(f"Target current last row by Column A: {target_last_row}")
        log(f"Paste start row: {paste_row}")
        log(f"Paste end row: {paste_row + row_count - 1}")

        # -----------------------------
        # Paste Data
        # -----------------------------
        log("Pasting Summary data to target sheet...")

        ws_target.Range(
            ws_target.Cells(paste_row, 1),
            ws_target.Cells(paste_row + row_count - 1, col_count)
        ).Value = data

        log("✓ Paste data completed")

        # -----------------------------
        # Stamp Run Date in Column R
        # -----------------------------
        log("Stamping Py Date in Column R...")

        ws_target.Cells(1, 18).Value = "Py Date"

        run_date = datetime.now().strftime("%d/%m/%Y")

        ws_target.Range(
            ws_target.Cells(paste_row, 18),
            ws_target.Cells(paste_row + row_count - 1, 18)
        ).Value = run_date

        ws_target.Columns(18).NumberFormat = "dd/mm/yyyy"

        log(f"✓ Py Date stamped at Column R: {run_date}")

        # -----------------------------
        # Header Column S
        # -----------------------------
        ws_target.Cells(1, 19).Value = "Restaurant Replace Log"

        # -----------------------------
        # Fill Formula After Append
        # -----------------------------
        log("Running Fill Formula after append...")

        fill_formulas_after_append(
            ws_target=ws_target,
            paste_row=paste_row,
            row_count=row_count,
            target_sheet_name=target_sheet_name
        )

        # -----------------------------
        # Read Report Week After Formula Fill
        # -----------------------------
        report_week = get_week_from_appended_rows(
            ws_target=ws_target,
            paste_row=paste_row,
            row_count=row_count,
            week_col_letter=WEEK_SOURCE_COLUMN
        )

        # -----------------------------
        # Read Latest Data Date From Appended Rows
        # -----------------------------
        latest_data_date = get_latest_date_from_appended_rows(
            ws_target=ws_target,
            paste_row=paste_row,
            row_count=row_count,
            date_col_letter=DATE_SOURCE_COLUMN
        )

        # Fallback: ถ้าคอลัมน์ A ไม่ใช่วันที่จริง ให้ดึงวันที่ล่าสุดจากชื่อชีท/ชื่อไฟล์ใน Buffe_VBA
        # เช่น SYY_DINNER_01-05-26 ถึง SYY_DINNER_17-05-26 => ใช้ 17/05/2026
        if latest_data_date is None:
            latest_data_date = get_latest_date_from_vba_sheet_names(vba_wb)

        # -----------------------------
        # Clean Restaurant Name + Log
        # -----------------------------
        log("Running Restaurant Name Clean...")

        clean_restaurant_names(
            ws_target=ws_target,
            target_sheet_name=target_sheet_name,
            paste_row=paste_row,
            row_count=row_count
        )

        # -----------------------------
        # Save Main File
        # -----------------------------
        log("Saving Main File...")
        main_wb.Save()

        log(
            f"✓ Append สำเร็จ: {target_sheet_name} | "
            f"Rows: {row_count} | Py Date: {run_date} | Week: {report_week}"
        )

        return row_count, report_week, latest_data_date

    finally:
        log("Closing Main File...")
        main_wb.Close(SaveChanges=True)
        log("✓ Main File closed")



# =====================================================
# 10) REFRESH PIVOT + FILTER LATEST 5 WEEKS
# =====================================================

def _week_number(text):
    if text is None:
        return None

    m = re.search(r"(\d+)", str(text))
    if not m:
        return None

    return int(m.group(1))


def _get_latest_week_items(pivot_field, top_n=5):
    items = []

    for item in pivot_field.PivotItems():
        item_name = str(item.Name).strip()
        week_no = _week_number(item_name)

        if week_no is not None:
            items.append((week_no, item_name))

    # ตัดซ้ำ แล้วเรียง Week มากไปน้อย
    unique = {}
    for week_no, item_name in items:
        unique[item_name] = week_no

    sorted_items = sorted(unique.items(), key=lambda x: x[1], reverse=True)
    latest_items = [item_name for item_name, week_no in sorted_items[:top_n]]

    return latest_items


def _find_week_field_from_pivot(pt):
    """
    หา PivotField ที่เป็น Week แบบยืดหยุ่น
    1) ชื่อ Field มีคำว่า Week
    2) PivotItems มี Pattern Week / ตัวเลขสัปดาห์
    """

    # รอบแรก: ชื่อ Field มีคำว่า week
    for pf in pt.PivotFields():
        try:
            pf_name = str(pf.Name).lower()
            if "week" in pf_name:
                return pf
        except Exception:
            pass

    # รอบสอง: ดูจาก Item ว่ามี Week number เยอะพอไหม
    for pf in pt.PivotFields():
        try:
            latest_items = _get_latest_week_items(pf, top_n=5)
            if len(latest_items) >= 2:
                return pf
        except Exception:
            pass

    return None


def filter_pivot_to_latest_5_weeks(ws, cell_address):
    """
    Refresh Pivot แล้ว Filter เฉพาะ 5 Week ล่าสุด
    cell_address คือ Cell ที่อยู่ใน Pivot / ช่อง Filter ที่ผู้ใช้ระบุ เช่น Z7, AA7, F5
    """

    log(f"Filter latest 5 weeks at {ws.Name}!{cell_address}")

    try:
        pt = ws.Range(cell_address).PivotTable
    except Exception as e:
        log(f"⚠ ไม่เจอ PivotTable ที่ {ws.Name}!{cell_address}: {e}")
        return

    try:
        pt.RefreshTable()
        log(f"✓ Refresh Pivot: {ws.Name} | {pt.Name}")
    except Exception as e:
        log(f"⚠ Refresh Pivot ไม่สำเร็จ: {ws.Name} | {pt.Name} | {e}")

    pf = None

    # ถ้า Cell นั้นจับ PivotField ได้ ให้ใช้ก่อน
    try:
        pf = ws.Range(cell_address).PivotField
        log(f"✓ PivotField from cell: {pf.Name}")
    except Exception:
        pf = _find_week_field_from_pivot(pt)
        if pf is not None:
            log(f"✓ PivotField auto detected: {pf.Name}")

    if pf is None:
        log(f"⚠ หา PivotField Week ไม่เจอที่ {ws.Name}!{cell_address}")
        return

    latest_items = _get_latest_week_items(pf, top_n=5)

    if not latest_items:
        log(f"⚠ ไม่พบ Week Items ใน PivotField: {pf.Name}")
        return

    log(f"Latest 5 Week Items: {latest_items}")

    try:
        # Page Field ให้เลือกหลายอันได้
        try:
            pf.EnableMultiplePageItems = True
        except Exception:
            pass

        # เปิด Week ล่าสุดก่อน กัน Error ที่ Excel ไม่ยอมให้ซ่อนทุก Item
        for item in pf.PivotItems():
            item_name = str(item.Name).strip()
            if item_name in latest_items:
                try:
                    item.Visible = True
                except Exception:
                    pass

        # ซ่อนที่ไม่ใช่ 5 Week ล่าสุด
        for item in pf.PivotItems():
            item_name = str(item.Name).strip()
            try:
                item.Visible = item_name in latest_items
            except Exception as item_error:
                log(f"⚠ Set visible ไม่ได้: {item_name} | {item_error}")

        log(f"✓ Filter Pivot สำเร็จ: {ws.Name}!{cell_address} | {latest_items}")

    except Exception as e:
        log(f"⚠ Filter Pivot ไม่สำเร็จ: {ws.Name}!{cell_address}: {e}")


def refresh_pivots_and_filter_latest_weeks(excel, main_file):
    """
    Phase สุดท้าย:
    - Refresh Pivot หน้า DB SNT, DB SKC, DB SPN, DB SYY
    - Filter เฉพาะ 5 Week ล่าสุด ตาม Cell ที่กำหนด
    """

    log("Opening Main File for Pivot Refresh + Latest Week Filter...")
    main_wb = excel.Workbooks.Open(normalize_excel_path(main_file))

    try:
        pivot_filter_map = {
            "DB SNT": ["Z7"],
            "DB SKC": ["AA7", "AA48"],
            "DB SPN": ["F5", "F50"],
            "DB SYY": ["AE8", "AE46", "AE83"],
        }

        for sheet_name, cell_list in pivot_filter_map.items():
            log("=" * 80)
            log(f"Refresh Pivot Sheet: {sheet_name}")

            try:
                ws = main_wb.Worksheets(sheet_name)
            except Exception:
                log(f"⚠ ไม่เจอชีท: {sheet_name}")
                continue

            # Refresh ทุก Pivot ในชีทก่อน
            try:
                for pt in ws.PivotTables():
                    pt.RefreshTable()
                    log(f"✓ Refresh PivotTable: {sheet_name} | {pt.Name}")
            except Exception as e:
                log(f"⚠ Refresh PivotTables ในชีท {sheet_name} ไม่สำเร็จ: {e}")

            # Filter เฉพาะตำแหน่งที่คุยกันไว้
            for cell_address in cell_list:
                filter_pivot_to_latest_5_weeks(ws, cell_address)

        log("Saving Main File after Pivot Refresh + Latest Week Filter...")
        main_wb.Save()
        log("✓ Pivot Refresh + Latest Week Filter completed")

    finally:
        log("Closing Main File after Pivot Refresh + Latest Week Filter...")
        main_wb.Close(SaveChanges=True)
        log("✓ Main File closed")


# =====================================================
# 11) SAVE AS BY LATEST DATA DATE
# =====================================================

def _normalize_date_year(dt):
    """ปรับปี พ.ศ. เป็น ค.ศ. และกันค่าที่ไม่ใช่วันที่งานจริงหลุดเข้ามา"""

    if dt is None:
        return None

    # ถ้าเจอปี พ.ศ. เช่น 2569 ให้แปลงเป็น 2026
    if dt.year >= 2400:
        try:
            dt = dt.replace(year=dt.year - 543)
        except Exception:
            return None

    # กันเคส Excel เอาตัวเลขอื่น / เวลา / serial แปลก ๆ มาแปลงเป็นปี 1900/2000 เช่น 110400
    # งานนี้เป็นข้อมูลปีปัจจุบัน/ใกล้เคียง จึงรับเฉพาะช่วงปีที่สมเหตุสมผล
    if dt.year < 2020 or dt.year > 2035:
        return None

    return dt


def _try_parse_excel_date(value):
    if value is None or str(value).strip() == "":
        return None

    if isinstance(value, datetime):
        return _normalize_date_year(value)

    # Excel serial date: รับเฉพาะ serial ที่แปลงแล้วอยู่ในช่วงปีที่สมเหตุสมผล
    if isinstance(value, (int, float)):
        try:
            dt = datetime(1899, 12, 30) + timedelta(days=float(value))
            return _normalize_date_year(dt)
        except Exception:
            return None

    text = str(value).strip()

    # ตัดเวลาออก ถ้าเป็น 17/05/2026 00:00:00
    text = re.sub(r"\s+\d{1,2}:\d{2}(:\d{2})?$", "", text).strip()

    formats = [
        "%d/%m/%Y", "%d/%m/%y",
        "%d-%m-%Y", "%d-%m-%y",
        "%Y-%m-%d",
    ]

    for fmt in formats:
        try:
            dt = datetime.strptime(text, fmt)
            return _normalize_date_year(dt)
        except Exception:
            pass

    return None



def _try_parse_date_from_text(text):
    """
    หา Date จากข้อความ เช่น ชื่อชีท/ชื่อไฟล์
    รองรับแพทเทิร์นประมาณ SYY_DINNER_17-05-26, SKC_ALL_170526, 17/05/2026
    ใช้เป็น fallback กรณีคอลัมน์ A ไม่ใช่วันที่จริง
    """

    if text is None:
        return None

    text = str(text).strip()

    patterns = [
        r"(?<!\d)(\d{1,2})[-_/\.](\d{1,2})[-_/\.](\d{2,4})(?!\d)",
        r"(?<!\d)(\d{2})(\d{2})(\d{2})(?!\d)",
    ]

    for pat in patterns:
        matches = re.findall(pat, text)

        for d, m, y in matches:
            try:
                d = int(d)
                m = int(m)
                y = int(y)

                if y < 100:
                    y += 2000

                dt = datetime(y, m, d)
                dt = _normalize_date_year(dt)

                if dt is not None:
                    return dt

            except Exception:
                continue

    return None


def get_latest_date_from_vba_sheet_names(vba_wb):
    """
    หา Date ล่าสุดจากชื่อชีทในไฟล์ Buffe_VBA
    เหมาะกับไฟล์ชื่อแบบ SYY_DINNER_01-05-26 / SYY_DINNER_17-05-26
    ใช้ fallback ตอนคอลัมน์ A ในข้อมูลที่ Append ไม่ใช่วันที่
    """

    latest_date = None
    latest_sheet = ""

    log("Fallback: Reading latest date from Buffe_VBA sheet names...")

    for ws in vba_wb.Worksheets:
        sheet_name = str(ws.Name).strip()

        if sheet_name.lower() in ["control", "summary"]:
            continue

        dt = _try_parse_date_from_text(sheet_name)

        if dt is None:
            continue

        if latest_date is None or dt > latest_date:
            latest_date = dt
            latest_sheet = sheet_name

    if latest_date:
        log(f"✓ Latest date from sheet name: {latest_date.strftime('%d/%m/%Y')} | Sheet: {latest_sheet}")
    else:
        log("⚠ หา Date จากชื่อชีท Buffe_VBA ไม่เจอ")

    return latest_date


def get_latest_date_from_appended_rows(ws_target, paste_row, row_count, date_col_letter=DATE_SOURCE_COLUMN):
    """
    หา 'วันที่สุดท้าย' จากข้อมูลชุดที่เพิ่ง Append เท่านั้น
    ใช้สำหรับ Save As ชื่อไฟล์ เช่น 04_Buffet All Resort 260126-170526.xlsx
    ไม่ไปไล่ทั้งไฟล์ เพราะอาจเจอค่าเก่าหรือค่าที่ไม่ใช่ Date แล้วทำให้ชื่อไฟล์เพี้ยน
    """

    date_col = ws_target.Range(f"{date_col_letter}1").Column
    first_row = paste_row
    last_row = paste_row + row_count - 1

    log(
        f"Reading Latest Data Date from appended rows: "
        f"{ws_target.Name}!{date_col_letter}{first_row}:{date_col_letter}{last_row}"
    )

    rng = ws_target.Range(
        ws_target.Cells(first_row, date_col),
        ws_target.Cells(last_row, date_col)
    )

    values = rng.Value

    if values is None:
        log("⚠ ไม่พบ Date ในข้อมูลที่ Append")
        return None

    if not isinstance(values[0], tuple):
        values = (values,)

    latest_date = None

    for row in values:
        dt = _try_parse_excel_date(row[0])
        if dt is None:
            continue

        if latest_date is None or dt > latest_date:
            latest_date = dt

    if latest_date:
        log(f"✓ Latest appended data date: {latest_date.strftime('%d/%m/%Y')}")
    else:
        log("⚠ หา Latest appended data date ไม่เจอ")

    return latest_date

def get_latest_date_from_data_sheets(excel, main_file):
    """
    หา Date ล่าสุดจากชีทข้อมูลที่ใส่ข้อมูลเข้าไป
    อิงคอลัมน์ I ของชีท Input หลัก
    """

    log("Opening Main File for latest date detection...")
    main_wb = excel.Workbooks.Open(normalize_excel_path(main_file))

    data_sheet_names = [
        "SPN Lunch", "SPN Dinner", "SPN All",
        "SYY All", "SYY Dinner",
        "SNT Dinner", "SNT All",
        "SKC All", "SKC Dinner",
    ]

    latest_date = None

    try:
        for sheet_name in data_sheet_names:
            try:
                ws = main_wb.Worksheets(sheet_name)
            except Exception:
                log(f"⚠ ไม่เจอชีทข้อมูล: {sheet_name}")
                continue

            date_col = ws.Range(f"{DATE_SOURCE_COLUMN}1").Column
            last_row = ws.Cells(ws.Rows.Count, date_col).End(XL_UP).Row

            if last_row < 2:
                log(f"⚠ {sheet_name} ไม่มีข้อมูล")
                continue

            rng = ws.Range(ws.Cells(2, date_col), ws.Cells(last_row, date_col))
            values = rng.Value

            if values is None:
                continue

            if not isinstance(values[0], tuple):
                values = (values,)

            for row in values:
                dt = _try_parse_excel_date(row[0])

                if dt is None:
                    continue

                if latest_date is None or dt > latest_date:
                    latest_date = dt

        if latest_date:
            log(f"✓ Latest data date detected: {latest_date.strftime('%d/%m/%Y')}")
        else:
            log("⚠ ไม่พบวันที่ล่าสุดจากชีทข้อมูล")

        return latest_date

    finally:
        main_wb.Close(SaveChanges=True)
        log("✓ Main File closed after latest date detection")


def save_main_file_as_latest_date(excel, main_file, latest_date_override=None):
    """
    Save As ไฟล์หลักโดยเปลี่ยนวันที่ด้านหลังชื่อไฟล์
    ตัวอย่าง:
    04_Buffet All Resort 260126-100526.xlsx
    -> 04_Buffet All Resort 260126-170526.xlsx

    เวอร์ชันแก้ล่าสุด:
    - Normalize path ก่อนส่งให้ Excel กัน Error C:\//Users/.../A9374610
    - ใช้ FileFormat ให้ตรงนามสกุลไฟล์
    - ถ้าวันที่ใหม่เท่ากับชื่อไฟล์เดิม จะ Save เฉย ๆ ไม่ SaveAs ทับตัวเอง
    - ถ้ามีไฟล์ปลายทางชื่อเดียวกันอยู่แล้ว จะลบก่อน SaveAs เพื่อกันชื่อชน
    """

    main_file = normalize_excel_path(main_file)

    if latest_date_override is not None:
        latest_date = latest_date_override
        log(f"✓ Use latest date from appended data for Save As: {latest_date.strftime('%d/%m/%Y')}")
    else:
        latest_date = get_latest_date_from_data_sheets(excel, main_file)

    if latest_date is None:
        log("⚠ ข้าม Save As เพราะหาวันที่ล่าสุดไม่เจอ")
        return main_file

    date_text = latest_date.strftime("%d%m%y")

    folder = os.path.dirname(main_file)
    file_name = os.path.basename(main_file)
    name, ext = os.path.splitext(file_name)

    # เปลี่ยนเฉพาะวันที่ท้ายชื่อ หลังขีด - เช่น 100526
    new_name = re.sub(r"-\d{6}$", f"-{date_text}", name)

    if new_name == name:
        new_name = f"{name}-{date_text}"

    new_file = normalize_excel_path(os.path.join(folder, new_name + ext))

    log(f"Current Main File: {main_file}")
    log(f"Save As target: {new_file}")

    # ถ้าเป็นไฟล์ชื่อเดิมอยู่แล้ว ไม่ต้อง SaveAs ซ้ำ เพราะ Excel อาจมองว่า Save ทับ workbook ที่เปิดอยู่
    if os.path.normcase(os.path.abspath(main_file)) == os.path.normcase(os.path.abspath(new_file)):
        log("✓ ชื่อไฟล์เป็นวันที่ล่าสุดอยู่แล้ว จึง Save เฉย ๆ ไม่ Save As")
        wb = excel.Workbooks.Open(Filename=main_file)
        try:
            wb.Save()
            log(f"✓ Save สำเร็จ: {main_file}")
        finally:
            wb.Close(SaveChanges=True)
            log("✓ Main File closed after Save")
        return main_file

    # ถ้ามีไฟล์ปลายทางอยู่แล้ว ให้ลบก่อน เพื่อไม่ให้ Excel SaveAs แล้วชนชื่อไฟล์เดิม
    if os.path.exists(new_file):
        try:
            os.remove(new_file)
            log(f"✓ Removed existing target file before Save As: {new_file}")
        except Exception as e:
            raise Exception(f"ลบไฟล์ปลายทางก่อน Save As ไม่ได้ อาจมีไฟล์เปิดอยู่: {new_file} | {e}")

    file_format = get_excel_file_format(new_file)

    wb = excel.Workbooks.Open(Filename=main_file)

    try:
        wb.SaveAs(Filename=new_file, FileFormat=file_format)
        log(f"✓ Save As สำเร็จ: {new_file}")
    finally:
        wb.Close(SaveChanges=True)
        log("✓ Main File closed after Save As")

    return new_file

# =====================================================
# 10) MAIN FLOW
# =====================================================

def main():
    global MAIN_FILE

    # STEP 0: เลือกไฟล์ Report หลักก่อนเริ่มรัน
    MAIN_FILE = select_main_file()

    # STEP 1: Backup ก่อนเปิด / แก้ไฟล์ใด ๆ
    backup_files_before_start()

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.ScreenUpdating = True  # เปิดไว้ให้เห็น Excel ทำงานจริงระหว่างรัน

    wb = None

    try:
        log("Opening VBA File...")
        wb = excel.Workbooks.Open(normalize_excel_path(VBA_FILE))
        ws_control = wb.Worksheets("Control")
        log("✓ VBA File opened")

        # ก่อนเริ่ม เช็กว่ามีชีทค้างไหม
        reset_if_needed(excel, wb)

        # อ่าน Control Rows
        control_rows = get_control_rows(ws_control)

        if not control_rows:
            log("ไม่พบข้อมูลใน Control")
            return

        log(f"พบงานทั้งหมด {len(control_rows)} รอบ")

        detected_report_week = None
        detected_latest_data_date = None

        # วนทีละรอบ
        for item in control_rows:
            row = item["row"]
            input_folder = item["input_folder"]
            branch = item["branch"]
            data_type = item["data_type"]

            print("=" * 100)
            log(f"เริ่มรอบ: Row {row} | {branch} | {data_type}")
            log(f"Folder: {input_folder}")

            try:
                # -----------------------------
                # Check Folder
                # -----------------------------
                log("Checking input folder...")

                if not os.path.isdir(input_folder):
                    update_status(ws_control, row, "Error", "ไม่พบโฟลเดอร์")
                    wb.Save()
                    log("✗ Error: ไม่พบโฟลเดอร์")
                    continue

                log("✓ Folder exists")

                file_count = count_xlsx_files(input_folder)

                if file_count == 0:
                    update_status(ws_control, row, "No File", "ไม่พบไฟล์ .xlsx ในโฟลเดอร์")
                    wb.Save()
                    log("⚠ No File: ไม่พบไฟล์ .xlsx")
                    continue

                # -----------------------------
                # Set Current Area
                # -----------------------------
                set_current_area(ws_control, input_folder, branch, data_type)

                update_status(
                    ws_control,
                    row,
                    "Running",
                    f"พบไฟล์ {file_count} ไฟล์"
                )

                wb.Save()

                # -----------------------------
                # STEP 1: Combine
                # -----------------------------
                log("STEP 1: Combine files to sheets")
                run_macro(excel, MACRO_COMBINE)
                wb.Save()

                # -----------------------------
                # STEP 2: Clean
                # -----------------------------
                log("STEP 2: Clean combined sheets")
                run_macro(excel, MACRO_CLEAN)
                wb.Save()

                # -----------------------------
                # STEP 3: Append Summary to Main File
                # -----------------------------
                log("STEP 3: Append Summary to Main File")

                appended_rows, report_week, latest_data_date = append_summary_to_main_file(
                    excel=excel,
                    vba_wb=wb,
                    main_file=MAIN_FILE,
                    branch=branch,
                    data_type=data_type
                )

                if report_week:
                    detected_report_week = report_week
                    log(f"✓ Update detected_report_week = {detected_report_week}")

                if latest_data_date:
                    if detected_latest_data_date is None or latest_data_date > detected_latest_data_date:
                        detected_latest_data_date = latest_data_date
                    log(f"✓ Update detected_latest_data_date = {detected_latest_data_date.strftime('%d/%m/%Y')}")

                update_status(
                    ws_control,
                    row,
                    "OK",
                    f"Append {appended_rows} rows | Files {file_count} | Week {report_week}"
                )

                wb.Save()

                # -----------------------------
                # STEP 4: Reset After Round
                # -----------------------------
                log("STEP 4: Reset after round")
                force_reset(excel, wb)

                log(f"✓ จบรอบ: {branch} | {data_type}")

            except Exception as e:
                error_msg = str(e)

                update_status(ws_control, row, "Error", error_msg[:250])
                wb.Save()

                log(f"✗ Error รอบ {branch} | {data_type}: {error_msg}")

                try:
                    force_reset(excel, wb)
                except Exception as reset_error:
                    log(f"✗ Reset หลัง Error ไม่สำเร็จ: {reset_error}")

                continue

        print("=" * 100)
        log("✓ ทำงานครบทุกแถวใน Control แล้ว")

        # -----------------------------
        # STEP 5: Update Table Week Dropdown
        # -----------------------------
        if detected_report_week:
            log(f"STEP 5: Update Table Week Dropdown by detected week: {detected_report_week}")

            update_table_week_dropdowns(
                excel=excel,
                main_file=MAIN_FILE,
                week_value=detected_report_week
            )
        else:
            log("⚠ ไม่พบ Report Week จากข้อมูลที่ Append จึงข้ามการ Update Dropdown")

        # -----------------------------
        # STEP 6: Refresh Pivot + Filter Latest 5 Weeks
        # -----------------------------
        log("STEP 6: Refresh Pivot + Filter Latest 5 Weeks")
        refresh_pivots_and_filter_latest_weeks(
            excel=excel,
            main_file=MAIN_FILE
        )

        # -----------------------------
        # STEP 7: Save As by Latest Data Date
        # -----------------------------
        log("STEP 7: Save As Main File by Latest Data Date")
        save_main_file_as_latest_date(
            excel=excel,
            main_file=MAIN_FILE,
            latest_date_override=detected_latest_data_date
        )

        wb.Save()

    except Exception as e:
        log(f"Main Error: {e}")

    finally:
        # ช่วงเทสยังเปิด Excel ทิ้งไว้ให้ดูผล
        # ถ้าต้องการให้ปิดเอง ค่อยเปิด 2 บรรทัดนี้
        # if wb is not None:
        #     wb.Close(SaveChanges=True)
        # excel.Quit()
        pass


if __name__ == "__main__":
    main()