import streamlit as st
import pandas as pd
import sqlite3
import io
import urllib.parse
import string
import os
from datetime import datetime, date, timedelta, time

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

# Set page configurations with native default theme formatting
st.set_page_config(page_title="Office Operations Portal", layout="wide")

LOGO1_PATH = "logo.png"
LOGO2_PATH = "logo2.png"

# ==============================================================================
# ⚙️ 1. HELPER FUNCTIONS & DATABASE ENGINE
# ==============================================================================
DB_FILE = "office_operations.db"

def generate_whatsapp_link(phone_number, text):
    """Generates a pre-filled WhatsApp click-to-chat URL."""
    clean_phone = phone_number.replace("+", "").replace(" ", "").replace("-", "") if phone_number else ""
    encoded_text = urllib.parse.quote(text)
    if clean_phone:
        return f"https://wa.me/{clean_phone}?text={encoded_text}"
    else:
        return f"https://api.whatsapp.com/send?text={encoded_text}"

def generate_car_name(index):
    """Generates sequential car names: Car A, Car B ... Car Z, Car AA, etc."""
    if index < 26:
        return f"Car {string.ascii_uppercase[index]}"
    else:
        first = string.ascii_uppercase[(index // 26) - 1]
        second = string.ascii_uppercase[index % 26]
        return f"Car {first}{second}"

def export_custom_batam_excel(groups_summary_df, detailed_df):
    """
    Generates a customized Excel workbook matching the 
    'Template - Daily Transportation Arrangement TUCC Batam.xlsx' design layout.
    """
    wb = openpyxl.Workbook()
    
    # --------------------------------------------------------------------------
    # SHEET 1: Summary Format
    # --------------------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "DAILY TRANSPORTATION"
    ws1.views.sheetView[0].showGridLines = True
    
    font_title = Font(name="Calibri", size=13, bold=True, color="1F4E78")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=10)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
    )
    
    # Merge B2:D4 for Header Block
    ws1.merge_cells("B2:D4")
    title_cell_s1 = ws1["B2"]
    title_cell_s1.value = "DAILY TRANSPORTATION ARRANGEMENT\nTUCC PROJECT - BATAM MODULE YARD [MD-1 & MD-4]\nJOB CODE : 0 - 0847 - 00 - 0001"
    title_cell_s1.font = font_title
    title_cell_s1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    if os.path.exists(LOGO1_PATH):
        try:
            img1 = OpenpyxlImage(LOGO1_PATH)
            img1.width = 110
            img1.height = 50
            ws1.add_image(img1, "A2")
        except Exception:
            pass

    if os.path.exists(LOGO2_PATH):
        try:
            img2 = OpenpyxlImage(LOGO2_PATH)
            img2.width = 130
            img2.height = 50
            ws1.add_image(img2, "H2")
        except Exception:
            pass

    headers = [
        "Car Group", "Vehicle Model", "Plate Number", "Color", "Driver Name", 
        "Contact Number", "Passenger(s)", "ETD 1 (From)", "ETD 2 (To)"
    ]
    
    start_row = 6
    for col_idx, h_title in enumerate(headers, start=1):
        cell = ws1.cell(row=start_row, column=col_idx, value=h_title)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        
    for r_idx, row_vals in enumerate(groups_summary_df.values, start=start_row+1):
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws1.cell(row=r_idx, column=c_idx, value="" if pd.isna(val) else val)
            cell.font = font_data
            cell.border = thin_border
            if c_idx in [1, 3, 4, 8, 9]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(vertical="center")

    for col in ws1.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws1.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # --------------------------------------------------------------------------
    # SHEET 2: DETAILED ALLOCATIONS (A3 Paper, Merged B2:D4, 3-Line Vehicle Desc)
    # --------------------------------------------------------------------------
    ws2 = wb.create_sheet(title="Detailed Allocations")
    ws2.views.sheetView[0].showGridLines = True
    
    # Configure Paper Size to A3
    ws2.page_setup.paperSize = ws2.PAPERSIZE_A3
    
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
        
        ws2.row_dimensions[r_idx].height = 42
        
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

    for col in ws2.columns:
        max_len = 0
        for cell in col:
            lines = str(cell.value or '').split('\n')
            for line in lines:
                if len(line) > max_len:
                    max_len = len(line)
        col_letter = col[0].column_letter
        ws2.column_dimensions[col_letter].width = max(max_len + 4, 18)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

def get_db_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. users table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT NOT NULL,
            role TEXT NOT NULL,
            email_recipients TEXT,
            emp_name TEXT
        )
    ''')
    
    # 2. holidays table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS holidays (
            holiday_date TEXT PRIMARY KEY,
            description TEXT
        )
    ''')
    
    # 3. overtime_requests table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS overtime_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT,
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
    
    # 4. meeting_rooms table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS meeting_rooms (
            room_number TEXT PRIMARY KEY,
            room_name TEXT,
            capacity INTEGER,
            location TEXT
        )
    ''')
    
    # 5. room_bookings table
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

    # 6. cars table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cars (
            car_name TEXT PRIMARY KEY,
            plate_number TEXT UNIQUE NOT NULL,
            vehicle TEXT,
            color TEXT DEFAULT 'Black'
        )
    ''')

    # 7. fleet_drivers table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS fleet_drivers (
            driver_name TEXT PRIMARY KEY,
            driver_mobile TEXT
        )
    ''')

    # 8. transit_groups table
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

    # 9. transit_passengers table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS transit_passengers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_name TEXT NOT NULL,
            passengers TEXT NOT NULL,
            FOREIGN KEY (group_name) REFERENCES transit_groups(group_name) ON DELETE CASCADE
        )
    ''')

    # 10. daily_transit table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS daily_transit (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            transit_date TEXT NOT NULL,
            group_name TEXT NOT NULL,
            FOREIGN KEY (group_name) REFERENCES transit_groups(group_name) ON DELETE CASCADE
        )
    ''')
    
    # Seed Admin user if not exists
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

# Schema migrations helper
def run_migrations():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("PRAGMA table_info(cars)")
    car_cols = [col[1] for col in cursor.fetchall()]
    if "color" not in car_cols:
        cursor.execute("ALTER TABLE cars ADD COLUMN color TEXT DEFAULT 'Black'")

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
    st.rerun()

tabs = ["⏰ Overtime & Transport", "👥 Transit Groups & Passengers", "📅 Daily Transit Dispatch", "🏢 Meeting Rooms", "🛠️ System Administration"]
tab1, tab1_b, tab1_c, tab2, tab3 = st.tabs(tabs)

# --- TAB 1: OVERTIME REQUESTS ---
with tab1:
    st.header("Request Overtime & Logistics Tracking")
    
    conn = get_db_connection()
    holidays_df = pd.read_sql_query("SELECT holiday_date FROM holidays", conn)
    holiday_list = holidays_df['holiday_date'].tolist()
    conn.close()
    
    col1, col2 = st.columns(2)
    with col1:
        ot_date = st.date_input("Select Target Date", value=date.today())
        date_str = ot_date.strftime("%Y-%m-%d")
        is_sunday = ot_date.weekday() == 6
        is_holiday = date_str in holiday_list
        
        if is_sunday or is_holiday:
            default_start = time(7, 0)
            default_end = time(15, 0)
            st.caption("ℹ️ Baseline Rule: **Sunday/Holiday (07:00 - 15:00)**.")
        else:
            default_start = time(17, 30)
            default_end = time(19, 0)
            st.caption("ℹ️ Baseline Rule: **Weekday/Saturday (17:30 - 19:00)**.")
        
        start_time = st.time_input("OT Start Time", value=default_start)
        end_time = st.time_input("OT End Time", value=default_end)
        needs_transport = st.selectbox("Require Individual Transportation Logistics?", ["Yes", "No"], index=0)
        
    with col2:
        if needs_transport == "Yes":
            if is_sunday or is_holiday:
                default_dep_time = start_time
                default_ret_time = end_time
                st.caption("ℹ️ Sunday/Holiday Rule Applied: Departure = Start Time, Return = End Time.")
            else:
                default_dep_time = end_time
                default_ret_time = time(23, 0)
            
            origin = st.text_input("Origin Address", value="Main Corporate Office")
            destination = st.text_input("Target Destination")
            dep_time = st.time_input("Departure Timeline Estimate", value=default_dep_time)
            ret_time = st.time_input("Return Timeline Estimate", value=default_ret_time)
        else:
            origin, destination, dep_time, ret_time = ["", "", "", ""]

    if st.button("Submit New Overtime Request"):
        conn = get_db_connection()
        conn.execute('''
            INSERT INTO overtime_requests (username, ot_date, start_time, end_time, needs_transport, 
            origin, destination, departure_time, return_time)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (st.session_state.username, date_str, start_time.strftime("%H:%M"), end_time.strftime("%H:%M"),
              needs_transport, origin, destination, str(dep_time), str(ret_time)))
        
        conn.commit()
        conn.close()
        st.success("🎉 Overtime log successfully submitted!")

    st.subheader("📋 Overtime Submission History Log")
    conn = get_db_connection()
    if st.session_state.role == "Admin":
        ot_df = pd.read_sql_query("SELECT * FROM overtime_requests", conn)
    else:
        ot_df = pd.read_sql_query("SELECT * FROM overtime_requests WHERE username = ?", conn, params=[st.session_state.username])
    conn.close()
    
    if not ot_df.empty:
        st.dataframe(ot_df, use_container_width=True)

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
        etd_1 = st.text_input("ETD 1 (From)", value="05:45")
        etd_2 = st.text_input("ETD 2 (To)", value="17:30")
        
        if st.button("Save Transit Group"):
            if selected_car_group != "No cars available" and selected_driver != "No drivers available":
                try:
                    conn = get_db_connection()
                    conn.execute("INSERT INTO transit_groups (group_name, driver_name, etd_1, etd_2) VALUES (?, ?, ?, ?)",
                                 (selected_car_group, selected_driver, etd_1, etd_2))
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
    st.subheader("📋 Configured Groups & Assigned Passengers")
    conn = get_db_connection()
    
    # Sheet 1 Data: Grouped Passengers
    groups_summary_df = pd.read_sql_query('''
        SELECT tg.group_name AS "Group Name", c.vehicle AS "Vehicle Model", c.plate_number AS "Plate Number", 
               c.color AS "Color", tg.driver_name AS "Driver Name", fd.driver_mobile AS "Contact Number", 
               GROUP_CONCAT(tp.passengers, ', ') AS "Passengers",
               tg.etd_1 AS "ETD 1 (From)", tg.etd_2 AS "ETD 2 (To)"
        FROM transit_groups tg
        LEFT JOIN cars c ON tg.group_name = c.car_name
        LEFT JOIN fleet_drivers fd ON tg.driver_name = fd.driver_name
        LEFT JOIN transit_passengers tp ON tg.group_name = tp.group_name
        GROUP BY tg.id
    ''', conn)

    # Sheet 2 Data: Unrolled Passengers for Detailed Allocations
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
    
    if not groups_summary_df.empty:
        st.dataframe(groups_summary_df, use_container_width=True)
        
        # Build Styled Batam TUCC Excel Workbook
        excel_bytes = export_custom_batam_excel(groups_summary_df, unrolled_df)

        st.download_button(
            label="📥 Download Custom Batam TUCC Excel Report (.xlsx)",
            data=excel_bytes,
            file_name=f"Daily_Transportation_Arrangement_TUCC_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- TAB 1C: DAILY TRANSIT DISPATCH ---
with tab1_c:
    st.header("📅 Daily Transit Dispatch Schedule")
    
    conn = get_db_connection()
    available_groups = [g['group_name'] for g in conn.execute("SELECT group_name FROM transit_groups").fetchall()]
    conn.close()
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Schedule Group for Date")
        dispatch_date = st.date_input("Select Transit Date", value=date.today())
        disp_date_str = dispatch_date.strftime("%Y-%m-%d")
        
        if available_groups:
            selected_dispatch_group = st.selectbox("Select Group Name to Dispatch", available_groups, key="disp_group_sel")
            
            if st.button("Schedule Daily Transit"):
                conn = get_db_connection()
                conn.execute("INSERT INTO daily_transit (transit_date, group_name) VALUES (?, ?)",
                             (disp_date_str, selected_dispatch_group))
                conn.commit()
                conn.close()
                st.success(f"Group '{selected_dispatch_group}' scheduled for {disp_date_str}!")
                st.rerun()
        else:
            st.warning("No Transit Groups created yet.")
            
    with col2:
        st.subheader("📢 Share Schedule via WhatsApp")
        filter_date = st.date_input("Filter Schedule Date", value=date.today(), key="filter_sched_date")
        filter_date_str = filter_date.strftime("%Y-%m-%d")
        
        conn = get_db_connection()
        daily_df = pd.read_sql_query('''
            SELECT dt.id AS daily_id, dt.transit_date, tg.group_name, c.vehicle, c.plate_number, c.color,
                   tg.driver_name, fd.driver_mobile, tg.etd_1, tg.etd_2, 
                   GROUP_CONCAT(tp.passengers, ', ') AS passengers
            FROM daily_transit dt
            JOIN transit_groups tg ON dt.group_name = tg.group_name
            LEFT JOIN cars c ON tg.group_name = c.car_name
            LEFT JOIN fleet_drivers fd ON tg.driver_name = fd.driver_name
            LEFT JOIN transit_passengers tp ON tg.group_name = tp.group_name
            WHERE dt.transit_date = ?
            GROUP BY dt.id
        ''', conn, params=[filter_date_str])
        conn.close()
        
        if not daily_df.empty:
            summary_text = f"機能 *TRANSPORTATION SUMMARY ({filter_date_str})*\n\n"
            
            for idx, row in daily_df.iterrows():
                summary_text += f"*Vehicle:* {row['vehicle'] or 'N/A'} (Color: {row['color'] or 'N/A'})\n"
                summary_text += f"*Plate Number:* {row['plate_number'] or 'N/A'}\n"
                summary_text += f"*Driver:* {row['driver_name'] or 'N/A'}\n"
                summary_text += f"*Driver Mobile No:* {row['driver_mobile'] or 'N/A'}\n"
                summary_text += f"*Passenger Name:* {row['passengers'] or 'N/A'}\n"
                summary_text += f"*From (etd_1):* {row['etd_1'] or 'N/A'}\n"
                summary_text += f"*To (etd_2):* {row['etd_2'] or 'N/A'}\n"
                summary_text += "-----------------------------------\n"
                
            wa_link = generate_whatsapp_link("", summary_text)
            st.link_button("📢 Send Transportation Summary to WhatsApp", wa_link)
            
            with st.expander("👁 Preview WhatsApp Summary Text"):
                st.text(summary_text)

    st.markdown("---")
    st.subheader(f"📊 Scheduled Dispatches for {disp_date_str}")
    if not daily_df.empty:
        st.dataframe(daily_df, use_container_width=True)
    else:
        st.info("No transit groups scheduled for this date.")

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
            recurrence_end = st.date_input("Recurrence End Target (Max 6 Months)", value=book_date + timedelta(days=7))
            
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
            column_config["group_name"] = st.column_config.SelectboxColumn("Group Name", options=groups_list)

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
