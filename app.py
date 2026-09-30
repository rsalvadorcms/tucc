import streamlit as st
import pandas as pd
import sqlite3
import io
import urllib.parse
import string
import os
import re
from datetime import datetime, date, timedelta, time

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

# Optional import for PDF rendering
try:
    from reportlab.lib.pagesizes import A3, portrait
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

# Set page configurations with native default theme formatting
st.set_page_config(page_title="Office Operations Portal", layout="wide")

LOGO1_PATH = "logo.png"
LOGO2_PATH = "logo2.png"

# ==============================================================================
# ⚙️ 1. HELPER FUNCTIONS & DATABASE ENGINE
# ==============================================================================
DB_FILE = "office_operations.db"

def format_military_time(input_str: str) -> str:
    """
    Cleans user input and enforces 24-hour military time HH:MM format.
    Accepts raw digits like '800' -> '08:00', '1730' -> '17:30', or '8:00' -> '08:00'.
    Returns None if the input is completely invalid.
    """
    if not input_str:
        return ""
    
    # Strip non-numeric characters
    clean_digits = re.sub(r'\D', '', str(input_str).strip())
    
    if len(clean_digits) == 3: # e.g. 800 -> 0800
        clean_digits = "0" + clean_digits
    elif len(clean_digits) == 1: # e.g. 8 -> 0800
        clean_digits = "0" + clean_digits + "00"
    elif len(clean_digits) == 2: # e.g. 17 -> 1700
        clean_digits = clean_digits + "00"
        
    if len(clean_digits) == 4:
        hours = int(clean_digits[:2])
        minutes = int(clean_digits[2:])
        if 0 <= hours <= 23 and 0 <= minutes <= 59:
            return f"{hours:02d}:{minutes:02d}"
            
    return None

def export_df_to_excel(df, sheet_name="Data"):
    """Generic helper function to export any pandas DataFrame to XLSX format."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()

def export_custom_batam_excel(detailed_df, effective_date_str=""):
    """
    Generates a customized Excel workbook containing ONLY the 'Detailed Allocations' sheet
    with repeating headers on print, row height = 20 for rows 2-4, white fill for A1:F5,
    fixed widths for Column B (12) and Column C (18), and Effective Date in E5:F5.
    """
    wb = openpyxl.Workbook()
    
    font_title = Font(name="Calibri", size=13, bold=True, color="1F4E78")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_bold_label = Font(name="Calibri", size=10, bold=True, color="000000")
    font_data = Font(name="Calibri", size=10)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
    )

    # DETAILED ALLOCATIONS SHEET
    ws2 = wb.active
    ws2.title = "Detailed Allocations"
    ws2.views.sheetView[0].showGridLines = True
    
    # Configure Paper Size to A3 Portrait and repeat rows 1 to 6 on every printed page
    ws2.page_setup.paperSize = ws2.PAPERSIZE_A3
    ws2.page_setup.orientation = ws2.ORIENTATION_PORTRAIT
    ws2.print_title_rows = '1:6'
    
    # Fill white color for cells A1:F5
    for r in range(1, 6):
        for c in range(1, 7):
            ws2.cell(row=r, column=c).fill = white_fill

    # Adjust row height to 20 for rows 2, 3, and 4
    ws2.row_dimensions[2].height = 20
    ws2.row_dimensions[3].height = 20
    ws2.row_dimensions[4].height = 20

    # Write Effective Date in E5 and F5
    cell_e5 = ws2.cell(row=5, column=5, value="Effective Date:")
    cell_e5.font = font_bold_label
    cell_e5.alignment = Alignment(horizontal="right", vertical="center")
    
    cell_f5 = ws2.cell(row=5, column=6, value=effective_date_str)
    cell_f5.font = font_bold_label
    cell_f5.alignment = Alignment(horizontal="center", vertical="center")

    start_row = 6
    
    # Merge cells B2:D4 for Title Header Block
    ws2.merge_cells("B2:D4")
    title_cell = ws2["B2"]
    title_cell.value = "Daily Transportation Arrangement - Passenger list\nTUCC PROJECT - BATAM MODULE YARD [MD-1 & MD-4]\nJOB CODE : 0 - 0847 - 00 - 0001"
    title_cell.font = font_title
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # Top-Left Logo (A2)
    if os.path.exists(LOGO1_PATH):
        try:
            img1_det = OpenpyxlImage(LOGO1_PATH)
            img1_det.width = 110
            img1_det.height = 50
            ws2.add_image(img1_det, "A2")
        except Exception:
            pass

    # Top-Right Logo (F2)
    if os.path.exists(LOGO2_PATH):
        try:
            img2_det = OpenpyxlImage(LOGO2_PATH)
            img2_det.width = 130
            img2_det.height = 50
            ws2.add_image(img2_det, "F2")
        except Exception:
            pass

    det_headers = [
        "Vehicle Description", "Driver Name", "Contact Number", 
        "Passenger", "ETD 1", "ETD 2"
    ]
    
    for col_idx, h_title in enumerate(det_headers, start=1):
        cell = ws2.cell(row=start_row, column=col_idx, value=h_title)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        
    data_start = start_row + 1
    
    for r_idx, row in enumerate(detailed_df.iterrows(), start=data_start):
        row_data = row[1]
        
        v_model = str(row_data.get("Vehicle Model", "") or "Standard Vehicle").strip()
        p_num = str(row_data.get("Plate Number", "") or "N/A").strip()
        v_color = str(row_data.get("Color", "") or "Black").strip()
        
        v_desc_3lines = f"{v_model}\n{p_num}\nColor : {v_color}"
            
        d_name = row_data.get("Driver Name", "")
        c_num = row_data.get("Contact Number", "")
        p_name = row_data.get("Passenger Name", "")
        etd1 = row_data.get("ETD 1 (From)", "")
        etd2 = row_data.get("ETD 2 (To)", "")
        
        ordered_vals = [v_desc_3lines, d_name, c_num, p_name, etd1, etd2]
        
        for c_idx, val in enumerate(ordered_vals, start=1):
            cell = ws2.cell(row=r_idx, column=c_idx, value="" if pd.isna(val) else val)
            cell.font = font_data
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Merge repeated vehicle & driver details vertically per group
    if not detailed_df.empty:
        current_grp = None
        grp_start = data_start
        tot_rows = len(detailed_df)
        
        for idx, grp_val in enumerate(detailed_df['Car Group'].values, start=data_start):
            if grp_val != current_grp:
                if current_grp is not None and (idx - 1) > grp_start:
                    for merge_col in [1, 2, 3, 5, 6]:
                        ws2.merge_cells(start_row=grp_start, start_column=merge_col, end_row=idx - 1, end_column=merge_col)
                        ws2.cell(row=grp_start, column=merge_col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                current_grp = grp_val
                grp_start = idx
                
        if current_grp is not None and (data_start + tot_rows - 1) > grp_start:
            for merge_col in [1, 2, 3, 5, 6]:
                ws2.merge_cells(start_row=grp_start, start_column=merge_col, end_row=data_start + tot_rows - 1, end_column=merge_col)
                ws2.cell(row=grp_start, column=merge_col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Apply column width logic
    for col in ws2.columns:
        col_letter = col[0].column_letter
        if col_letter == 'B':
            ws2.column_dimensions['B'].width = 12
        elif col_letter == 'C':
            ws2.column_dimensions['C'].width = 18
        else:
            max_len = 0
            for cell in col:
                lines = str(cell.value or '').split('\n')
                for line in lines:
                    if len(line) > max_len:
                        max_len = len(line)
            ws2.column_dimensions[col_letter].width = max(max_len + 4, 18)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

def export_custom_batam_pdf(detailed_df, effective_date_str=""):
    """
    Generates a PDF on A3 Portrait matching the Excel layout with Effective Date in E5/F5 block.
    """
    if not HAS_REPORTLAB:
        return None
        
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=portrait(A3),
        rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25
    )
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold',
        fontSize=15, leading=19, alignment=1, textColor=colors.HexColor('#1F4E78')
    )
    subtitle_style = ParagraphStyle(
        'DocSubtitle', parent=styles['Normal'], fontName='Helvetica-Bold',
        fontSize=11, leading=14, alignment=1, textColor=colors.HexColor('#333333')
    )
    eff_date_style = ParagraphStyle(
        'EffDateStyle', parent=styles['Normal'], fontName='Helvetica-Bold',
        fontSize=10, leading=12, alignment=2, textColor=colors.HexColor('#000000')
    )
    cell_style = ParagraphStyle(
        'CellText', parent=styles['Normal'], fontName='Helvetica',
        fontSize=9, leading=11, alignment=1
    )
    header_style = ParagraphStyle(
        'HeaderStyle', parent=styles['Normal'], fontName='Helvetica-Bold',
        fontSize=10, leading=12, alignment=1, textColor=colors.white
    )
    
    story = []
    
    title_p = Paragraph("Daily Transportation Arrangement - Passenger list", title_style)
    sub1_p = Paragraph("TUCC PROJECT - BATAM MODULE YARD [MD-1 & MD-4]", subtitle_style)
    sub2_p = Paragraph("JOB CODE : 0 - 0847 - 00 - 0001", subtitle_style)
    
    header_box = [title_p, Spacer(1, 3), sub1_p, Spacer(1, 2), sub2_p]
    
    img1_elem = RLImage(LOGO1_PATH, width=100, height=45) if os.path.exists(LOGO1_PATH) else ""
    img2_elem = RLImage(LOGO2_PATH, width=120, height=45) if os.path.exists(LOGO2_PATH) else ""
    
    top_table_data = [[img1_elem, header_box, img2_elem]]
    top_table = Table(top_table_data, colWidths=[110, 560, 110])
    top_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
        ('ALIGN', (2, 0), (2, 0), 'RIGHT'),
    ]))
    
    story.append(top_table)
    story.append(Spacer(1, 10))
    
    eff_p = Paragraph(f"<b>Effective Date:</b> {effective_date_str}", eff_date_style)
    eff_table = Table([[Paragraph("", cell_style), eff_p]], colWidths=[550, 230])
    eff_table.setStyle(TableStyle([
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(eff_table)
    story.append(Spacer(1, 8))
    
    headers = ["Vehicle Description", "Driver Name", "Contact Number", "Passenger", "ETD 1", "ETD 2"]
    table_data = [[Paragraph(h, header_style) for h in headers]]
    
    table_spans = []
    current_grp = None
    grp_start = 1
    
    for idx, row_data in enumerate(detailed_df.iterrows(), start=1):
        r_val = row_data[1]
        grp_name = r_val.get("Car Group", "")
        
        v_model = str(r_val.get("Vehicle Model", "") or "Standard Vehicle").strip()
        p_num = str(r_val.get("Plate Number", "") or "N/A").strip()
        v_color = str(r_val.get("Color", "") or "Black").strip()
        v_desc_html = f"{v_model}<br/>{p_num}<br/>Color : {v_color}"
        
        row_cells = [
            Paragraph(v_desc_html, cell_style),
            Paragraph(str(r_val.get("Driver Name", "") or ""), cell_style),
            Paragraph(str(r_val.get("Contact Number", "") or ""), cell_style),
            Paragraph(str(r_val.get("Passenger Name", "") or ""), cell_style),
            Paragraph(str(r_val.get("ETD 1 (From)", "") or ""), cell_style),
            Paragraph(str(r_val.get("ETD 2 (To)", "") or ""), cell_style)
        ]
        table_data.append(row_cells)
        
        if grp_name != current_grp:
            if current_grp is not None and (idx - 1) > grp_start:
                for col_i in [0, 1, 2, 4, 5]:
                    table_spans.append(('SPAN', (col_i, grp_start), (col_i, idx - 1)))
            current_grp = grp_name
            grp_start = idx

    tot_rows = len(detailed_df)
    if current_grp is not None and tot_rows > grp_start:
        for col_i in [0, 1, 2, 4, 5]:
            table_spans.append(('SPAN', (col_i, grp_start), (col_i, tot_rows)))
            
    main_table = Table(table_data, colWidths=[150, 120, 120, 210, 90, 90], repeatRows=1)
    
    ts = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BFBFBF')),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ] + table_spans
    
    main_table.setStyle(TableStyle(ts))
    story.append(main_table)
    doc.build(story)
    
    buffer.seek(0)
    return buffer.getvalue()

def get_db_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT NOT NULL,
            role TEXT NOT NULL,
            email_recipients TEXT,
            emp_name TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS holidays (
            holiday_date TEXT PRIMARY KEY,
            description TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS overtime_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT,
            emp_name TEXT,
            ot_date TEXT,
            start_time TEXT,
            end_time TEXT,
            needs_transport TEXT DEFAULT 'Yes',
            origin TEXT,
            destination TEXT,
            departure_time TEXT,
            return_time TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS meeting_rooms (
            room_number TEXT PRIMARY KEY,
            room_name TEXT,
            capacity INTEGER,
            location TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS room_bookings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            room_number TEXT,
            booked_by TEXT,
            booking_date TEXT,
            start_time TEXT,
            end_time TEXT,
            is_recurring TEXT,
            recurrence_end_date TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cars (
            car_name TEXT PRIMARY KEY,
            plate_number TEXT UNIQUE NOT NULL,
            vehicle TEXT,
            color TEXT DEFAULT 'Black'
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS fleet_drivers (
            driver_name TEXT PRIMARY KEY,
            driver_mobile TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS transit_groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_name TEXT UNIQUE NOT NULL,
            driver_name TEXT,
            etd_1 TEXT,
            etd_2 TEXT,
            FOREIGN KEY (group_name) REFERENCES cars(car_name) ON DELETE CASCADE,
            FOREIGN KEY (driver_name) REFERENCES fleet_drivers(driver_name) ON DELETE SET NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS transit_passengers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_name TEXT NOT NULL,
            passengers TEXT NOT NULL,
            FOREIGN KEY (group_name) REFERENCES transit_groups(group_name) ON DELETE CASCADE
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS daily_transit (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            transit_date_start TEXT NOT NULL,
            transit_date_end TEXT NOT NULL,
            group_name TEXT DEFAULT 'TBA',
            requested_by TEXT,
            etd_1 TEXT,
            etd_2 TEXT,
            location_from TEXT,
            location_to TEXT,
            daily TEXT DEFAULT 'No'
        )
    ''')
    
    cursor.execute("SELECT * FROM users WHERE username='admin'")
    if not cursor.fetchone():
        cursor.execute("INSERT INTO users VALUES ('admin', 'admin123', 'Admin', 'admin@company.com', 'System Administrator')")
        
    cursor.execute("SELECT COUNT(*) FROM meeting_rooms")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO meeting_rooms VALUES ('101', 'Boardroom', 15, '1st Floor')")
        cursor.execute("INSERT INTO meeting_rooms VALUES ('102', 'Huddle Room Alpha', 6, '2nd Floor')")

    cursor.execute("SELECT COUNT(*) FROM fleet_drivers")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO fleet_drivers VALUES ('John Doe', '+628111222333')")
        cursor.execute("INSERT INTO fleet_drivers VALUES ('Jane Smith', '+628999888777')")

    cursor.execute("SELECT COUNT(*) FROM cars")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO cars VALUES ('Car A', 'B 1234 ABC', 'Toyota Avanza', 'Black')")
        cursor.execute("INSERT INTO cars VALUES ('Car B', 'B 5678 XYZ', 'Toyota Innova', 'White')")
        
    conn.commit()
    conn.close()

init_db()

def run_migrations():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("PRAGMA table_info(cars)")
    car_cols = [col[1] for col in cursor.fetchall()]
    if "color" not in car_cols:
        cursor.execute("ALTER TABLE cars ADD COLUMN color TEXT DEFAULT 'Black'")

    cursor.execute("PRAGMA table_info(overtime_requests)")
    ot_cols = [col[1] for col in cursor.fetchall()]
    if "emp_name" not in ot_cols:
        cursor.execute("ALTER TABLE overtime_requests ADD COLUMN emp_name TEXT")

    # Migrations for daily_transit table
    cursor.execute("PRAGMA table_info(daily_transit)")
    dt_cols = [col[1] for col in cursor.fetchall()]

    if "transit_date" in dt_cols and "transit_date_start" not in dt_cols:
        cursor.execute("ALTER TABLE daily_transit RENAME COLUMN transit_date TO transit_date_start")
    
    cursor.execute("PRAGMA table_info(daily_transit)")
    dt_cols_updated = [col[1] for col in cursor.fetchall()]
    if "transit_date_end" not in dt_cols_updated:
        cursor.execute("ALTER TABLE daily_transit ADD COLUMN transit_date_end TEXT")
        cursor.execute("UPDATE daily_transit SET transit_date_end = transit_date_start WHERE transit_date_end IS NULL")

    new_dt_fields = {
        "requested_by": "TEXT",
        "etd_1": "TEXT",
        "etd_2": "TEXT",
        "location_from": "TEXT",
        "location_to": "TEXT",
        "daily": "TEXT DEFAULT 'No'"
    }
    for col_name, col_type in new_dt_fields.items():
        if col_name not in dt_cols_updated:
            cursor.execute(f"ALTER TABLE daily_transit ADD COLUMN {col_name} {col_type}")

    conn.commit()
    conn.close()

run_migrations()

# ==============================================================================
# 🔐 2. AUTHENTICATION USER INTERFACE
# ==============================================================================
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""
    st.session_state.emp_name = ""

if 'emp_name' not in st.session_state:
    st.session_state.emp_name = st.session_state.get('username', '')

if not st.session_state.logged_in:
    st.title("🏢 Office Operations Portal")
    st.subheader("Login to access scheduling & overtime systems")
    
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log In")
        
        if submitted:
            conn = get_db_connection()
            user = conn.execute("SELECT * FROM users WHERE username=? AND password=?", (username, password)).fetchone()
            conn.close()
            if user:
                st.session_state.logged_in = True
                st.session_state.username = user['username']
                st.session_state.role = user['role']
                st.session_state.emp_name = user['emp_name'] or user['username']
                st.rerun()
            else:
                st.error("Invalid username or password configuration.")
    st.stop()

# ==============================================================================
# 🗂️ 3. MAIN APP CONTROL PANELS
# ==============================================================================
st.sidebar.title(f"👋 Welcome, {st.session_state.username}")
st.sidebar.info(f"Access Level: **{st.session_state.role}**")
if st.sidebar.button("Logout Profile"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""
    st.session_state.emp_name = ""
    st.rerun()

tabs = ["⏰ Overtime & Transport", "👥 Transit Groups & Passengers", "📅 Daily Transit Dispatch", "🏢 Meeting Rooms", "🛠️ System Administration"]
tab1, tab1_b, tab1_c, tab2, tab3 = st.tabs(tabs)

# --- TAB 1: OVERTIME REQUESTS ---
with tab1:
    st.header("Request Overtime & Logistics Tracking")
    
    conn = get_db_connection()
    holidays_df = pd.read_sql_query("SELECT holiday_date FROM holidays", conn)
    holiday_list = holidays_df['holiday_date'].tolist()
    
    users_df = pd.read_sql_query("SELECT emp_name FROM users WHERE emp_name IS NOT NULL AND emp_name != ''", conn)
    all_emp_names = users_df['emp_name'].tolist()
    conn.close()
    
    current_user_emp = st.session_state.get("emp_name", "") or st.session_state.get("username", "")

    if not all_emp_names:
        all_emp_names = [current_user_emp] if current_user_emp else ["Default Employee"]
        
    logged_in_emp = current_user_emp if current_user_emp in all_emp_names else all_emp_names[0]

    col1, col2 = st.columns(2)
    with col1:
        selected_staff_members = st.multiselect(
            "Select Staff Member(s) for Overtime", 
            options=all_emp_names, 
            default=[logged_in_emp]
        )
        
        ot_date = st.date_input("Select Target Date", value=date.today(), key="ot_date_picker")
        date_str = ot_date.strftime("%Y-%m-%d") if ot_date else ""
        
        is_sunday = ot_date.weekday() == 6 if ot_date else False
        is_holiday = date_str in holiday_list if date_str else False
        
        if is_sunday or is_holiday:
            default_start = time(7, 0)
            default_end = time(15, 0)
            default_origin = "Panbil"
            default_dest = "Yard-1 Office"
            default_dep_time_str = "07:00"
            st.caption("ℹ️ Baseline Rule: **Sunday/Holiday (07:00 - 15:00)**.")
        else:
            default_start = time(17, 30)
            default_end = time(19, 0)
            default_origin = "Yard-1 Office"
            default_dest = "Panbil"
            default_dep_time_str = "19:00"
            st.caption("ℹ️ Baseline Rule: **Weekday/Saturday (17:30 - 19:00)**.")
        
        start_time = st.time_input("OT Start Time (24-hr Military Time)", value=default_start)
        end_time = st.time_input("OT End Time (24-hr Military Time)", value=default_end)
        needs_transport = st.selectbox("Require Individual Transportation Logistics?", ["Yes", "No"], index=0)
        
    with col2:
        if needs_transport == "Yes":
            origin = st.text_input("Origin Address", value=default_origin)
            destination = st.text_input("Target Destination", value=default_dest)
            dep_time_str = st.text_input("Departure Timeline Estimate (24-hr)", value=default_dep_time_str)
            ret_time_str = st.text_input("Return Timeline Estimate (24-hr)", value="")
        else:
            origin, destination, dep_time_str, ret_time_str = ["", "", "", ""]

    if st.button("Submit New Overtime Request"):
        if not ot_date or not date_str:
            st.error("❌ Submission Failed: Target Date cannot be blank.")
        elif start_time is None:
            st.error("❌ Submission Failed: OT Start Time cannot be blank.")
        elif end_time is None:
            st.error("❌ Submission Failed: OT End Time cannot be blank.")
        elif not selected_staff_members:
            st.error("❌ Submission Failed: Please select at least one staff member.")
        else:
            start_time_military = start_time.strftime("%H:%M")
            end_time_military = end_time.strftime("%H:%M")
            
            conn = get_db_connection()
            for staff in selected_staff_members:
                conn.execute('''
                    INSERT INTO overtime_requests (username, emp_name, ot_date, start_time, end_time, needs_transport, 
                    origin, destination, departure_time, return_time)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (st.session_state.username, staff, date_str, start_time_military, end_time_military,
                      needs_transport, origin, destination, dep_time_str, ret_time_str))
            
            conn.commit()
            conn.close()
            st.success(f"🎉 Overtime log successfully submitted for {len(selected_staff_members)} staff member(s)!")
            st.rerun()

    st.subheader("📋 Overtime Submission History Log")
    conn = get_db_connection()
    if st.session_state.role == "Admin":
        ot_df = pd.read_sql_query("SELECT id, username, emp_name AS 'Employee Name', ot_date AS 'Date', start_time AS 'Start Time', end_time AS 'End Time', needs_transport AS 'Needs Transport', origin AS 'Origin', destination AS 'Destination', departure_time AS 'Departure Time', return_time AS 'Return Time' FROM overtime_requests", conn)
    else:
        ot_df = pd.read_sql_query("SELECT id, username, emp_name AS 'Employee Name', ot_date AS 'Date', start_time AS 'Start Time', end_time AS 'End Time', needs_transport AS 'Needs Transport', origin AS 'Origin', destination AS 'Destination', departure_time AS 'Departure Time', return_time AS 'Return Time' FROM overtime_requests WHERE username = ?", conn, params=[st.session_state.username])
    conn.close()
    
    if not ot_df.empty:
        st.dataframe(ot_df, use_container_width=True)
        
        ot_excel_bytes = export_df_to_excel(ot_df, sheet_name="Overtime_Requests")
        st.download_button(
            label="📥 Export Overtime Log to Excel (.xlsx)",
            data=ot_excel_bytes,
            file_name=f"Overtime_Requests_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- TAB 1B: TRANSIT GROUPS & PASSENGERS MANAGEMENT ---
with tab1_b:
    st.header("👥 Transit Groups & Passengers Management")
    
    conn = get_db_connection()
    drivers_list = [d['driver_name'] for d in conn.execute("SELECT driver_name FROM fleet_drivers").fetchall()]
    cars_list = [c['car_name'] for c in conn.execute("SELECT car_name FROM cars").fetchall()]
    employees_list = [u['emp_name'] for u in conn.execute("SELECT emp_name FROM users WHERE emp_name IS NOT NULL AND emp_name != ''").fetchall()]
    groups_list = [g['group_name'] for g in conn.execute("SELECT group_name FROM transit_groups").fetchall()]
    conn.close()

    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1. Create Transit Group")
        selected_car_group = st.selectbox("Group Name (Select from Cars)", cars_list if cars_list else ["No cars available"], key="g_car_select")
        selected_driver = st.selectbox("Assign Driver (from Fleet Drivers)", drivers_list if drivers_list else ["No drivers available"])
        
        raw_etd1_grp = st.text_input("ETD 1 (From) [e.g. 0545 or 05:45]", value="05:45", key="grp_etd1_in")
        raw_etd2_grp = st.text_input("ETD 2 (To) [e.g. 1730 or 17:30]", value="17:30", key="grp_etd2_in")
        
        if st.button("Save Transit Group"):
            etd1_grp_formatted = format_military_time(raw_etd1_grp)
            etd2_grp_formatted = format_military_time(raw_etd2_grp)
            
            if not etd1_grp_formatted or not etd2_grp_formatted:
                st.error("❌ Invalid Time Format: Please enter valid 24-hr military time (e.g. 0545 or 1730).")
            elif selected_car_group != "No cars available" and selected_driver != "No drivers available":
                try:
                    conn = get_db_connection()
                    conn.execute("INSERT INTO transit_groups (group_name, driver_name, etd_1, etd_2) VALUES (?, ?, ?, ?)",
                                 (selected_car_group, selected_driver, etd1_grp_formatted, etd2_grp_formatted))
                    conn.commit()
                    conn.close()
                    st.success(f"Group '{selected_car_group}' created successfully!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("A group for this car already exists.")
            else:
                st.error("Please ensure a valid Car and Driver are selected.")

    with col2:
        st.subheader("2. Assign Passengers to Group")
        if groups_list and employees_list:
            selected_group_for_p = st.selectbox("Select Group Name (from Transit Groups)", groups_list, key="p_group_select")
            selected_passengers = st.multiselect("Select Passenger(s) (from User Employee Names)", employees_list)
            
            if st.button("Assign Passengers to Group"):
                if selected_passengers:
                    conn = get_db_connection()
                    for passenger in selected_passengers:
                        conn.execute("INSERT INTO transit_passengers (group_name, passengers) VALUES (?, ?)",
                                     (selected_group_for_p, passenger))
                    conn.commit()
                    conn.close()
                    st.success("Passengers added successfully!")
                    st.rerun()
                else:
                    st.error("Please select at least one passenger.")
        else:
            st.info("Ensure Transit Groups are created and Users have 'emp_name' populated.")

    st.markdown("---")
    
    title_col, eff_date_col = st.columns([2, 1])
    with title_col:
        st.subheader("📋 Configured Groups & Assigned Passengers")
    with eff_date_col:
        target_effective_date = st.date_input("Target Effective Date", value=date.today(), key="eff_date_picker")
        eff_date_str = target_effective_date.strftime("%Y-%m-%d") if target_effective_date else ""

    conn = get_db_connection()
    unrolled_df = pd.read_sql_query('''
        SELECT tg.group_name AS "Car Group", c.vehicle AS "Vehicle Model", c.plate_number AS "Plate Number", 
               c.color AS "Color", tg.driver_name AS "Driver Name", fd.driver_mobile AS "Contact Number", 
               tg.etd_1 AS "ETD 1 (From)", tg.etd_2 AS "ETD 2 (To)", tp.passengers AS "Passenger Name"
        FROM transit_groups tg
        LEFT JOIN cars c ON tg.group_name = c.car_name
        LEFT JOIN fleet_drivers fd ON tg.driver_name = fd.driver_name
        LEFT JOIN transit_passengers tp ON tg.group_name = tp.group_name
        ORDER BY tg.group_name
    ''', conn)
    conn.close()
    
    if not unrolled_df.empty:
        st.dataframe(unrolled_df, use_container_width=True)
        
        excel_bytes = export_custom_batam_excel(unrolled_df, effective_date_str=eff_date_str)
        pdf_bytes = export_custom_batam_pdf(unrolled_df, effective_date_str=eff_date_str) if HAS_REPORTLAB else None

        col_ex, col_pdf = st.columns(2)
        
        with col_ex:
            st.download_button(
                label="📥 Export to Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"Daily_Transportation_Arrangement_TUCC_{datetime.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        with col_pdf:
            if HAS_REPORTLAB and pdf_bytes:
                st.download_button(
                    label="📄 Export to PDF (.pdf)",
                    data=pdf_bytes,
                    file_name=f"Daily_Transportation_Arrangement_TUCC_{datetime.today().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            else:
                st.info("📄 PDF Export: Add `reportlab` to `requirements.txt` to enable.")

# --- TAB 1C: DAILY TRANSIT DISPATCH ---
with tab1_c:
    st.header("📅 Daily Transit Dispatch Schedule & Route Setting")
    
    conn = get_db_connection()
    users_df = pd.read_sql_query("SELECT emp_name FROM users WHERE emp_name IS NOT NULL AND emp_name != ''", conn)
    all_emp_names = users_df['emp_name'].tolist()
    
    cars_db_df = pd.read_sql_query("SELECT car_name, plate_number, vehicle FROM cars", conn)
    conn.close()
    
    current_user_emp = st.session_state.get("emp_name", "") or st.session_state.get("username", "")
    if not all_emp_names:
        all_emp_names = [current_user_emp] if current_user_emp else ["Default Employee"]
    
    default_req_by = current_user_emp if current_user_emp in all_emp_names else all_emp_names[0]
    is_admin = st.session_state.get("role", "") == "Admin"

    # Build Car Options Dictionary with Plate Numbers
    car_option_labels = ["TBA - To Be Assigned"]
    car_label_to_group = {"TBA - To Be Assigned": "TBA"}
    car_group_to_label = {"TBA": "TBA - To Be Assigned"}
    
    if not cars_db_df.empty:
        for _, c_row in cars_db_df.iterrows():
            c_name = str(c_row['car_name']).strip()
            p_num = str(c_row['plate_number'] or 'N/A').strip()
            v_model = str(c_row['vehicle'] or '').strip()
            
            label_str = f"{c_name} - {p_num}" + (f" ({v_model})" if v_model else "")
            car_option_labels.append(label_str)
            car_label_to_group[label_str] = c_name
            car_group_to_label[c_name] = label_str

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Configure Transit Request")
        
        requested_by = st.selectbox("Requested By", options=all_emp_names, index=all_emp_names.index(default_req_by))
        
        if is_admin:
            is_daily = st.selectbox("Daily / Recurring Journey?", options=["No", "Yes"], index=0)
        else:
            is_daily = st.selectbox("Daily / Recurring Journey?", options=["No"], index=0, disabled=True, help="Recurring journey requests require Administrator permissions.")

        min_date = date.today()
        max_date = min_date + timedelta(days=30)
        
        dispatch_date_start = st.date_input("Select Transit Start Date", value=min_date, min_value=min_date, max_value=max_date, key="reg_dt_start")
        disp_date_start_str = dispatch_date_start.strftime("%Y-%m-%d") if dispatch_date_start else ""
        
        if is_daily == "Yes":
            dispatch_date_end = st.date_input(
                "Select Transit End Date", 
                value=dispatch_date_start + timedelta(days=1), 
                min_value=dispatch_date_start + timedelta(days=1),
                max_value=max_date + timedelta(days=180),
                key="reg_dt_end"
            )
            disp_date_end_str = dispatch_date_end.strftime("%Y-%m-%d") if dispatch_date_end else ""
            st.caption("ℹ️ **Recurring Daily Journey**: Start Date must be earlier than End Date.")
        else:
            dispatch_date_end = dispatch_date_start
            disp_date_end_str = disp_date_start_str
            st.caption("ℹ️ Single journey request: **Transit End Date automatically set to Start Date**.")
        
        location_from = st.text_input("Origin Location (Location From)", value="Yard-1 Office")
        location_to = st.text_input("Target Location (Location To)", value="Yard-3 Office")

    with col2:
        st.subheader("Schedule & Vehicle Allocation")
        
        # Strict Military Time inputs (Auto-formatting support)
        raw_etd1 = st.text_input("ETD 1 (Start Time) [e.g. 0800 or 08:00]", value="08:00", placeholder="0800 or 08:00")
        raw_etd2 = st.text_input("ETD 2 (Return Time) [e.g. 1700 or 17:00]", value="17:00", placeholder="1700 or 17:00")
        
        if is_admin:
            selected_car_label = st.selectbox("Assigned Group / Car Name (With Plate Number)", options=car_option_labels, index=0)
            selected_group = car_label_to_group.get(selected_car_label, "TBA")
        else:
            selected_car_label = st.selectbox("Assigned Group / Car Name (With Plate Number)", options=["TBA - To Be Assigned"], index=0, disabled=True, help="Group name assignment is managed by Administrator.")
            selected_group = "TBA"

    if st.button("Submit Transit Dispatch Request"):
        formatted_etd1 = format_military_time(raw_etd1)
        formatted_etd2 = format_military_time(raw_etd2)
        
        if not formatted_etd1:
            st.error("❌ Invalid ETD 1 Time Format: Please enter a valid military time (e.g., 0800 or 08:00).")
        elif not formatted_etd2:
            st.error("❌ Invalid ETD 2 Time Format: Please enter a valid military time (e.g., 1700 or 17:00).")
        elif not disp_date_start_str or not disp_date_end_str:
            st.error("❌ Submission Failed: Transit Start and End Dates cannot be blank.")
        elif is_daily == "Yes" and dispatch_date_start >= dispatch_date_end:
            st.error("❌ Submission Failed: For daily recurring journeys, Start Date must be strictly earlier than End Date.")
        elif not location_from or not location_to:
            st.error("❌ Submission Failed: Origin and Target Locations cannot be blank.")
        else:
            conn = get_db_connection()
            conn.execute('''
                INSERT INTO daily_transit (transit_date_start, transit_date_end, group_name, requested_by, etd_1, etd_2, location_from, location_to, daily)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (disp_date_start_str, disp_date_end_str, selected_group, requested_by, formatted_etd1, formatted_etd2, location_from, location_to, is_daily))
            conn.commit()
            conn.close()
            st.success(f"🎉 Transit dispatch successfully requested for {requested_by} ({disp_date_start_str} to {disp_date_end_str})!")
            st.rerun()

    st.markdown("---")
    st.subheader("📊 Scheduled Transit Dispatches Log")
    
    conn = get_db_connection()
    daily_raw_df = pd.read_sql_query('''
        SELECT dt.id, dt.transit_date_start, dt.transit_date_end, dt.requested_by, 
               dt.group_name, c.plate_number, c.vehicle,
               dt.location_from, dt.location_to, dt.etd_1, dt.etd_2, dt.daily
        FROM daily_transit dt
        LEFT JOIN cars c ON dt.group_name = c.car_name
        ORDER BY dt.transit_date_start DESC, dt.id DESC
    ''', conn)
    conn.close()
    
    if not daily_raw_df.empty:
        display_df = daily_raw_df.copy()
        
        display_df['Group / Car'] = display_df.apply(
            lambda r: f"{r['group_name']} - {r['plate_number']}" if pd.notna(r['plate_number']) and r['group_name'] != 'TBA' else r['group_name'],
            axis=1
        )
        
        export_df = display_df[[
            "id", "transit_date_start", "transit_date_end", "requested_by", 
            "Group / Car", "location_from", "location_to", "etd_1", "etd_2", "daily"
        ]].rename(columns={
            "id": "Dispatch ID",
            "transit_date_start": "Transit Date Start",
            "transit_date_end": "Transit Date End",
            "requested_by": "Requested By",
            "location_from": "Origin (From)",
            "location_to": "Destination (To)",
            "etd_1": "ETD Start",
            "etd_2": "ETD Return",
            "daily": "Daily Recurring"
        })
        
        st.dataframe(export_df, use_container_width=True)
        
        dispatch_excel_bytes = export_df_to_excel(export_df, sheet_name="Daily_Dispatches")
        st.download_button(
            label="📥 Export Daily Dispatch Log to Excel (.xlsx)",
            data=dispatch_excel_bytes,
            file_name=f"Daily_Dispatch_Schedule_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # --- ADMIN MANAGEMENT & SEARCH FILTER CONSOLE ---
        if is_admin:
            st.markdown("---")
            st.subheader("🛠️ Admin Management: Edit / Delete Dispatch Records")
            
            with st.expander("🔍 Search & Filter Dispatch Records", expanded=True):
                sf_col1, sf_col2, sf_col3 = st.columns([2, 2, 2])
                
                with sf_col1:
                    search_query = st.text_input("Search (Requester, Origin, Dest, Car, Plate)", value="", key="search_dispatch_txt")
                
                with sf_col2:
                    filter_car_label = st.selectbox("Filter by Group / Car", options=["All"] + car_option_labels, index=0, key="filter_grp_sel")
                    
                with sf_col3:
                    filter_date_range = st.date_input(
                        "Filter by Date Range", 
                        value=(date.today() - timedelta(days=7), date.today() + timedelta(days=30)),
                        key="filter_dt_range"
                    )

            filtered_df = daily_raw_df.copy()
            
            if search_query:
                sq = search_query.lower()
                filtered_df = filtered_df[
                    filtered_df['requested_by'].astype(str).str.lower().str.contains(sq) |
                    filtered_df['location_from'].astype(str).str.lower().str.contains(sq) |
                    filtered_df['location_to'].astype(str).str.lower().str.contains(sq) |
                    filtered_df['group_name'].astype(str).str.lower().str.contains(sq) |
                    filtered_df['plate_number'].astype(str).str.lower().str.contains(sq)
                ]
                
            if filter_car_label != "All":
                target_grp_code = car_label_to_group.get(filter_car_label, "TBA")
                filtered_df = filtered_df[filtered_df['group_name'] == target_grp_code]
                
            if isinstance(filter_date_range, tuple) and len(filter_date_range) == 2:
                start_f, end_f = filter_date_range
                filtered_df['dt_start_obj'] = pd.to_datetime(filtered_df['transit_date_start']).dt.date
                filtered_df = filtered_df[(filtered_df['dt_start_obj'] >= start_f) & (filtered_df['dt_start_obj'] <= end_f)]

            st.caption(f"Showing **{len(filtered_df)}** of **{len(daily_raw_df)}** recorded dispatches.")
            
            if not filtered_df.empty:
                for idx, row in filtered_df.iterrows():
                    rec_id = row['id']
                    grp_disp = f"{row['group_name']} - {row['plate_number']}" if pd.notna(row['plate_number']) and row['group_name'] != 'TBA' else row['group_name']
                    rec_title = f"ID #{rec_id} | {row['transit_date_start']} ➡️ {row['transit_date_end']} | {row['requested_by']} | {row['location_from']} ➡️ {row['location_to']} ({grp_disp})"
                    
                    with st.expander(f"✏️ Manage Record: {rec_title}"):
                        e_col1, e_col2 = st.columns(2)
                        
                        with e_col1:
                            try:
                                curr_start_obj = datetime.strptime(row['transit_date_start'], "%Y-%m-%d").date()
                            except (ValueError, TypeError):
                                curr_start_obj = date.today()

                            try:
                                curr_end_obj = datetime.strptime(row['transit_date_end'], "%Y-%m-%d").date()
                            except (ValueError, TypeError):
                                curr_end_obj = curr_start_obj
                                
                            edit_req_by = st.selectbox("Requested By", options=all_emp_names, index=all_emp_names.index(row['requested_by']) if row['requested_by'] in all_emp_names else 0, key=f"e_req_{rec_id}")
                            edit_daily = st.selectbox("Daily Recurring?", options=["No", "Yes"], index=["No", "Yes"].index(row['daily'] if row['daily'] in ["Yes", "No"] else "No"), key=f"e_daily_{rec_id}")
                            
                            edit_dt_start = st.date_input("Transit Date Start", value=curr_start_obj, key=f"e_dt_start_{rec_id}")
                            
                            if edit_daily == "Yes":
                                edit_dt_end = st.date_input("Transit Date End", value=max(curr_end_obj, edit_dt_start + timedelta(days=1)), min_value=edit_dt_start + timedelta(days=1), key=f"e_dt_end_{rec_id}")
                            else:
                                edit_dt_end = edit_dt_start
                                st.caption("ℹ️ Non-recurring: Date End automatically set to Date Start.")

                            edit_loc_from = st.text_input("Origin Location", value=row['location_from'] or "Yard-1 Office", key=f"e_loc_from_{rec_id}")
                            edit_loc_to = st.text_input("Target Location", value=row['location_to'] or "Yard-3 Office", key=f"e_loc_to_{rec_id}")

                        with e_col2:
                            edit_raw_etd1 = st.text_input("ETD 1 (Start Time)", value=row['etd_1'] or "08:00", key=f"e_etd1_{rec_id}")
                            edit_raw_etd2 = st.text_input("ETD 2 (Return Time)", value=row['etd_2'] or "17:00", key=f"e_etd2_{rec_id}")
                            
                            curr_grp_code = row['group_name'] if row['group_name'] in car_group_to_label else "TBA"
                            curr_car_label = car_group_to_label.get(curr_grp_code, "TBA - To Be Assigned")
                            
                            edit_car_label = st.selectbox("Group / Car Name (With Plate Number)", options=car_option_labels, index=car_option_labels.index(curr_car_label) if curr_car_label in car_option_labels else 0, key=f"e_grp_{rec_id}")
                            edit_grp = car_label_to_group.get(edit_car_label, "TBA")

                        btn_col1, btn_col2 = st.columns([1, 4])
                        
                        with btn_col1:
                            if st.button("💾 Save Changes", key=f"btn_save_{rec_id}"):
                                edit_fmt_etd1 = format_military_time(edit_raw_etd1)
                                edit_fmt_etd2 = format_military_time(edit_raw_etd2)
                                
                                if not edit_fmt_etd1:
                                    st.error("❌ Invalid ETD 1 Time Format: Please enter a valid military time (e.g., 0800 or 08:00).")
                                elif not edit_fmt_etd2:
                                    st.error("❌ Invalid ETD 2 Time Format: Please enter a valid military time (e.g., 1700 or 17:00).")
                                else:
                                    conn = get_db_connection()
                                    conn.execute('''
                                        UPDATE daily_transit
                                        SET transit_date_start=?, transit_date_end=?, group_name=?, requested_by=?, etd_1=?, etd_2=?, location_from=?, location_to=?, daily=?
                                        WHERE id=?
                                    ''', (edit_dt_start.strftime("%Y-%m-%d"), edit_dt_end.strftime("%Y-%m-%d"), edit_grp, edit_req_by, edit_fmt_etd1, edit_fmt_etd2, edit_loc_from, edit_loc_to, edit_daily, rec_id))
                                    conn.commit()
                                    conn.close()
                                    st.success(f"Record ID #{rec_id} updated successfully!")
                                    st.rerun()

                        with btn_col2:
                            if st.button("🗑️ Delete Record", key=f"btn_del_{rec_id}", type="primary"):
                                conn = get_db_connection()
                                conn.execute("DELETE FROM daily_transit WHERE id=?", (rec_id,))
                                conn.commit()
                                conn.close()
                                st.success(f"Record ID #{rec_id} deleted!")
                                st.rerun()
            else:
                st.info("No dispatch records match your search filter criteria.")
    else:
        st.info("No transit dispatches scheduled yet.")

# --- TAB 2: MEETING ROOM BOOKINGS ENGINE ---
with tab2:
    st.header("Meeting Space Reservations Desk")
    
    conn = get_db_connection()
    rooms = conn.execute("SELECT * FROM meeting_rooms").fetchall()
    conn.close()
    
    if not rooms:
        st.warning("No physical boardrooms are registered yet.")
    else:
        room_options = {f"{r['room_name']} (Room {r['room_number']} - Capacity: {r['capacity']})": r['room_number'] for r in rooms}
        selected_room_label = st.selectbox("Choose Target Room Venue", list(room_options.keys()))
        selected_room_num = room_options[selected_room_label]
        
        col1, col2 = st.columns(2)
        with col1:
            book_date = st.date_input("Reservation Date", value=date.today(), key="bk_date")
            b_start = st.time_input("Reservation Start Time", value=time(9, 0), key="bk_start")
            b_end = st.time_input("Reservation End Time", value=time(10, 0), key="bk_end")
        
        with col2:
            recurrence = st.selectbox("Recurrence Schedule Pattern", ["None", "Daily", "Weekly", "Monthly"])
            max_rec_end = date.today() + timedelta(days=180) 
            recurrence_end = st.date_input("Recurrence End Target (Max 6 Months)", value=book_date + timedelta(days=7), key="rec_end_picker")
            
            if recurrence_end > max_rec_end:
                st.error("⚠️ Max 6 months recurrence limit exceeded.")
                st.stop()

        if st.button("Confirm Room Block Assignment"):
            target_dates = [book_date]
            if recurrence != "None":
                current_date = book_date
                while True:
                    if recurrence == "Daily":
                        current_date += timedelta(days=1)
                    elif recurrence == "Weekly":
                        current_date += timedelta(weeks=1)
                    elif recurrence == "Monthly":
                        current_date += timedelta(days=30)
                    
                    if current_date <= recurrence_end:
                        target_dates.append(current_date)
                    else:
                        break
            
            conflict_detected = False
            conn = get_db_connection()
            
            for t_date in target_dates:
                t_date_str = t_date.strftime("%Y-%m-%d")
                conflicts = conn.execute('''
                    SELECT * FROM room_bookings 
                    WHERE room_number = ? 
                    AND booking_date = ? 
                    AND NOT (start_time >= ? OR end_time <= ?)
                ''', (selected_room_num, t_date_str, b_end.strftime("%H:%M"), b_start.strftime("%H:%M"))).fetchall()
                
                if conflicts:
                    st.error(f"❌ Schedule Collision on {t_date_str}!")
                    conflict_detected = True
                    break
            
            if not conflict_detected:
                for t_date in target_dates:
                    conn.execute('''
                        INSERT INTO room_bookings (room_number, booked_by, booking_date, start_time, end_time, is_recurring, recurrence_end_date)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (selected_room_num, st.session_state.username, t_date.strftime("%Y-%m-%d"), 
                          b_start.strftime("%H:%M"), b_end.strftime("%H:%M"), recurrence, recurrence_end.strftime("%Y-%m-%d")))
                
                conn.commit()
                st.success(f"🎉 Room reserved across {len(target_dates)} intervals!")
            conn.close()

    st.subheader("📊 Master Room Allocation Schedules")
    conn = get_db_connection()
    bookings_df = pd.read_sql_query('''
        SELECT b.id, r.room_name, b.room_number, b.booked_by, b.booking_date, b.start_time, b.end_time, b.is_recurring 
        FROM room_bookings b 
        JOIN meeting_rooms r ON b.room_number = r.room_number
    ''', conn)
    conn.close()
    
    if not bookings_df.empty:
        st.dataframe(bookings_df, use_container_width=True)
        
        rooms_excel_bytes = export_df_to_excel(bookings_df, sheet_name="Room_Bookings")
        st.download_button(
            label="📥 Export Room Bookings to Excel (.xlsx)",
            data=rooms_excel_bytes,
            file_name=f"Room_Bookings_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- TAB 3: SYSTEM MASTER ADMINISTRATION CONTROL BOARDS ---
with tab3:
    if st.session_state.role != "Admin":
        st.error("🛡️ Restricted Access Control: Admin clearance required.")
    else:
        st.header("Admin Control Dashboard Engine")
        
        # --- EXCEL DATA IMPORT SECTION ---
        st.markdown("---")
        st.subheader("📤 Bulk Import Data via Excel (.xlsx)")
        
        import_table = st.selectbox("Select Database Table to Import Data Into", ["users", "holidays", "fleet_drivers", "cars", "transit_passengers"], key="import_tbl_sel")
        
        table_schemas = {
            "users": ["username", "password", "role", "email_recipients", "emp_name"],
            "holidays": ["holiday_date", "description"],
            "fleet_drivers": ["driver_name", "driver_mobile"],
            "cars": ["car_name", "plate_number", "vehicle", "color"],
            "transit_passengers": ["group_name", "passengers"]
        }
        
        req_cols = table_schemas[import_table]
        st.caption(f"ℹ️ **Required Excel (.xlsx) Headers for `{import_table}`:** `{', '.join(req_cols)}`")
        
        buffer_template = io.BytesIO()
        template_df = pd.DataFrame(columns=req_cols)
        with pd.ExcelWriter(buffer_template, engine='openpyxl') as writer:
            template_df.to_excel(writer, index=False, sheet_name=f"{import_table}_Template")
        excel_template_bytes = buffer_template.getvalue()
        
        st.download_button(
            label=f"📥 Download Excel (.xlsx) Template for {import_table}",
            data=excel_template_bytes,
            file_name=f"{import_table}_import_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        uploaded_file = st.file_uploader(f"Upload Excel file (.xlsx) for '{import_table}'", type=["xlsx", "xls"])
        
        if uploaded_file is not None:
            try:
                import_df = pd.read_excel(uploaded_file)
                
                import_df.columns = [str(c).strip().lower() for c in import_df.columns]
                missing_cols = [c for c in req_cols if c not in import_df.columns]
                
                if missing_cols:
                    st.error(f"❌ File missing required column headers: `{', '.join(missing_cols)}`")
                else:
                    st.write("🔍 **Preview Import Excel Data:**")
                    st.dataframe(import_df[req_cols], use_container_width=True)
                    
                    if st.button(f"🚀 Import {len(import_df)} Records into '{import_table}'"):
                        conn = get_db_connection()
                        cursor = conn.cursor()
                        
                        placeholders = ", ".join(["?"] * len(req_cols))
                        cols_str = ", ".join(req_cols)
                        
                        success_count = 0
                        for _, row in import_df.iterrows():
                            vals = [None if pd.isna(row[c]) else str(row[c]).strip() for c in req_cols]
                            cursor.execute(f"INSERT OR REPLACE INTO {import_table} ({cols_str}) VALUES ({placeholders})", vals)
                            success_count += 1
                            
                        conn.commit()
                        conn.close()
                        st.success(f"🎉 Successfully imported/updated {success_count} records in `{import_table}`!")
                        st.rerun()
            except Exception as e:
                st.error(f"Error processing file: {str(e)}")

        # --- INLINE CRUD DATA EDITOR ---
        st.markdown("---")
        st.subheader("🗃️ Master Data Tables Inline CRUD Editor")
        table_options = ["users", "holidays", "overtime_requests", "meeting_rooms", "room_bookings", "fleet_drivers", "cars", "transit_groups", "transit_passengers", "daily_transit"]
        selected_table = st.selectbox("Choose Database Table to Manage", table_options)
        
        conn = get_db_connection()
        table_df = pd.read_sql_query(f"SELECT * FROM {selected_table}", conn)
        
        original_passwords = {}
        if selected_table == "users":
            original_passwords = dict(zip(table_df["username"], table_df["password"]))
            table_df["password"] = "••••••••"

        drivers_list = [d['driver_name'] for d in conn.execute("SELECT driver_name FROM fleet_drivers").fetchall()]
        cars_list = [c['car_name'] for c in conn.execute("SELECT car_name FROM cars").fetchall()]
        groups_list = [g['group_name'] for g in conn.execute("SELECT group_name FROM transit_groups").fetchall()]
        emp_list = [u['emp_name'] for u in conn.execute("SELECT emp_name FROM users WHERE emp_name IS NOT NULL AND emp_name != ''").fetchall()]
        conn.close()
        
        column_config = {}
        if selected_table == "overtime_requests":
            column_config["needs_transport"] = st.column_config.SelectboxColumn("Needs Transport", options=["Yes", "No"])
        elif selected_table == "transit_groups":
            column_config["group_name"] = st.column_config.SelectboxColumn("Group Name (Car)", options=cars_list)
            column_config["driver_name"] = st.column_config.SelectboxColumn("Driver Name", options=drivers_list)
        elif selected_table == "transit_passengers":
            column_config["group_name"] = st.column_config.SelectboxColumn("Group Name", options=groups_list)
            column_config["passengers"] = st.column_config.SelectboxColumn("Passenger", options=emp_list)
        elif selected_table == "daily_transit":
            column_config["group_name"] = st.column_config.SelectboxColumn("Group Name", options=["TBA"] + groups_list)
            column_config["daily"] = st.column_config.SelectboxColumn("Daily Recurring", options=["Yes", "No"])

        st.markdown("💡 *Edit cells or use dropdowns where configured. Passwords are masked.*")
        
        edited_df = st.data_editor(table_df, num_rows="dynamic", use_container_width=True, column_config=column_config, key=f"editor_{selected_table}")
        
        if st.button(f"Save Grid Changes to Database ({selected_table})"):
            conn = get_db_connection()
            cursor = conn.cursor()
            
            cursor.execute(f"DELETE FROM {selected_table}")
            
            for _, row in edited_df.iterrows():
                row_dict = row.to_dict()
                
                if selected_table == "users":
                    u_name = row_dict.get("username")
                    if row_dict.get("password") == "••••••••":
                        row_dict["password"] = original_passwords.get(u_name, "")
                
                columns = [k for k in row_dict.keys() if row_dict[k] is not None]
                values = [row_dict[k] for k in columns]
                placeholders = ", ".join(["?"] * len(columns))
                col_names = ", ".join(columns)
                
                cursor.execute(f"INSERT INTO {selected_table} ({col_names}) VALUES ({placeholders})", values)
                
            conn.commit()
            conn.close()
            st.success(f"🎉 Database '{selected_table}' updated successfully!")
            st.rerun()
