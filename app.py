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

def format_excel_worksheet(ws, df, title_text, logo_path="logo.png", merge_repeat_cols=None):
    """Applies corporate styling, colors, logo insertion, auto column width, and cell merging to an openpyxl worksheet."""
    
    # Styles definition
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Dark Blue
    zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")  # Very Light Blue
    title_font = Font(name="Arial", size=14, bold=True, color="1F4E78")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    start_row = 1
    
    # 1. Add Logo if file exists
    if os.path.exists(logo_path):
        try:
            img = OpenpyxlImage(logo_path)
            img.width = 120
            img.height = 50
            ws.add_image(img, "A1")
            start_row = 5 # Push table down if logo exists
        except Exception:
            start_row = 1

    # 2. Add Worksheet Title
    title_cell = ws.cell(row=start_row, column=1, value=title_text)
    title_cell.font = title_font
    start_row += 2

    # 3. Write Header
    headers = list(df.columns)
    for col_num, header_title in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=col_num, value=header_title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        
    data_start_row = start_row + 1

    # 4. Write Data Rows
    for row_idx, row_data in enumerate(df.values, start=data_start_row):
        is_even = (row_idx % 2 == 0)
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value="" if pd.isna(value) else value)
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")
            if is_even:
                cell.fill = zebra_fill

    # 5. Merge Repeated Cells across groups (if requested)
    if merge_repeat_cols and not df.empty:
        current_group = None
        group_start_row = data_start_row
        
        for r_idx, row_val in enumerate(df['group_name'].values, start=data_start_row):
            if row_val != current_group:
                # Merge previous group range
                if current_group is not None and (r_idx - 1) > group_start_row:
                    for col_c in merge_repeat_cols:
                        ws.merge_cells(start_row=group_start_row, start_column=col_c, end_row=r_idx - 1, end_column=col_c)
                        merged_cell = ws.cell(row=group_start_row, column=col_c)
                        merged_cell.alignment = Alignment(horizontal="center", vertical="center")
                current_group = row_val
                group_start_row = r_idx
        
        # Merge final group
        if current_group is not None and (data_start_row + len(df) - 1) > group_start_row:
            for col_c in merge_repeat_cols:
                ws.merge_cells(start_row=group_start_row, start_column=col_c, end_row=data_start_row + len(df) - 1, end_column=col_c)
                merged_cell = ws.cell(row=group_start_row, column=col_c)
                merged_cell.alignment = Alignment(horizontal="center", vertical="center")

    # 6. Auto-fit Column Widths
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

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
            vehicle TEXT
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
        cursor.execute("INSERT INTO cars VALUES ('Car A', 'B 1234 ABC', 'Toyota Avanza')")
        cursor.execute("INSERT INTO cars VALUES ('Car B', 'B 5678 XYZ', 'Toyota Innova')")
        
    conn.commit()
    conn.close()

init_db()

# Schema migrations helper
def run_migrations():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Check fleet_drivers table
    cursor.execute("PRAGMA table_info(fleet_drivers)")
    fd_cols = [col[1] for col in cursor.fetchall()]
    
    if "plate_number" in fd_cols:
        cursor.execute("SELECT DISTINCT plate_number, vehicle FROM fleet_drivers WHERE plate_number IS NOT NULL AND plate_number != ''")
        existing_cars = cursor.fetchall()
        
        for idx, car in enumerate(existing_cars):
            c_name = generate_car_name(idx)
            p_num = car['plate_number']
            v_type = car['vehicle'] if 'vehicle' in fd_cols and car['vehicle'] else 'Standard Vehicle'
            cursor.execute("INSERT OR IGNORE INTO cars (car_name, plate_number, vehicle) VALUES (?, ?, ?)", (c_name, p_num, v_type))
            
        cursor.execute('''
            CREATE TABLE fleet_drivers_new (
                driver_name TEXT PRIMARY KEY,
                driver_mobile TEXT
            )
        ''')
        cursor.execute("INSERT INTO fleet_drivers_new (driver_name, driver_mobile) SELECT driver_name, driver_mobile FROM fleet_drivers")
        cursor.execute("DROP TABLE fleet_drivers")
        cursor.execute("ALTER TABLE fleet_drivers_new RENAME TO fleet_drivers")

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
    
    # Sheet 1 Data: Grouped Passengers (Comma-separated)
    groups_df = pd.read_sql_query('''
        SELECT tg.group_name AS "Group Name", c.vehicle AS "Vehicle", c.plate_number AS "Plate Number", 
               tg.driver_name AS "Driver Name", fd.driver_mobile AS "Driver Mobile", 
               tg.etd_1 AS "ETD 1", tg.etd_2 AS "ETD 2", GROUP_CONCAT(tp.passengers, ', ') AS "Passengers"
        FROM transit_groups tg
        LEFT JOIN cars c ON tg.group_name = c.car_name
        LEFT JOIN fleet_drivers fd ON tg.driver_name = fd.driver_name
        LEFT JOIN transit_passengers tp ON tg.group_name = tp.group_name
        GROUP BY tg.id
    ''', conn)

    # Sheet 2 Data: Unrolled Passengers (Individual rows per passenger)
    unrolled_df = pd.read_sql_query('''
        SELECT tg.group_name AS "group_name", tg.group_name AS "Group Name", c.vehicle AS "Vehicle", c.plate_number AS "Plate Number", 
               tg.driver_name AS "Driver Name", fd.driver_mobile AS "Driver Mobile", 
               tg.etd_1 AS "ETD 1", tg.etd_2 AS "ETD 2", tp.passengers AS "Passenger Name"
        FROM transit_groups tg
        LEFT JOIN cars c ON tg.group_name = c.car_name
        LEFT JOIN fleet_drivers fd ON tg.driver_name = fd.driver_name
        LEFT JOIN transit_passengers tp ON tg.group_name = tp.group_name
        ORDER BY tg.group_name
    ''', conn)
    conn.close()
    
    if not groups_df.empty:
        st.dataframe(groups_df, use_container_width=True)
        
        # Build Styled Multi-Sheet Workbook with Cell Merging & Formatting
        wb = openpyxl.Workbook()
        
        # Sheet 1: Grouped
        ws1 = wb.active
        ws1.title = "Grouped Summary"
        format_excel_worksheet(ws1, groups_df, title_text="TRANSIT GROUPS SUMMARY REPORT")
        
        # Sheet 2: Detailed Unrolled with Merged Group Columns
        ws2 = wb.create_sheet(title="Detailed Passengers")
        # Prepare display dataframe (omit sorting helper column 'group_name')
        unrolled_display_df = unrolled_df.drop(columns=["group_name"])
        # Columns 1 to 7 (Group Name, Vehicle, Plate Number, Driver, Mobile, ETD 1, ETD 2) will be merged per group
        merge_cols = [1, 2, 3, 4, 5, 6, 7]
        format_excel_worksheet(ws2, unrolled_display_df, title_text="DETAILED PASSENGER ALLOCATIONS REPORT", merge_repeat_cols=merge_cols)
        
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_bytes = excel_buffer.getvalue()

        st.download_button(
            label="📥 Download Formatted Excel Report (.xlsx)",
            data=excel_bytes,
            file_name=f"Configured_Transit_Groups_{datetime.today().strftime('%Y%m%d')}.xlsx",
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
            SELECT dt.id AS daily_id, dt.transit_date, tg.group_name, c.vehicle, c.plate_number,
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
                summary_text += f"*Vehicle:* {row['vehicle'] or 'N/A'}\n"
                summary_text += f"*Plate Number:* {row['plate_number'] or 'N/A'}\n"
                summary_text += f"*Driver:* {row['driver_name'] or 'N/A'}\n"
                summary_text += f"*Driver Mobile No:* {row['driver_mobile'] or 'N/A'}\n"
                summary_text += f"*Passenger Name:* {row['passengers'] or 'N/A'}\n"
                summary_text += f"*From (etd_1):* {row['etd_1'] or 'N/A'}\n"
                summary_text += f"*To (etd_2):* {row['etd_2'] or 'N/A'}\n"
                summary_text += "-----------------------------------\n"
                
            wa_link = generate_whatsapp_link("", summary_text)
            st.link_button("📢 Send Transportation Summary to WhatsApp", wa_link)
            
            with st.expander("👁️ Preview WhatsApp Summary Text"):
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
            "cars": ["car_name", "plate_number", "vehicle"],
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
        
        # Keep track of original passwords for 'users' table
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
