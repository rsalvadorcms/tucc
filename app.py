import streamlit as st
import pandas as pd
import sqlite3
import io
from datetime import datetime, date, timedelta, time

# Set page configurations with native default theme formatting
st.set_page_config(page_title="Office Operations Portal", layout="wide")

# ==============================================================================
# ⚙️ 1. SELF-CONTAINED DATABASE ENGINE
# ==============================================================================
DB_FILE = "office_operations.db"

def get_db_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Create Users Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT NOT NULL,
            role TEXT NOT NULL,
            email_recipients TEXT
        )
    ''')
    
    # Create Holidays Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS holidays (
            holiday_date TEXT PRIMARY KEY,
            description TEXT
        )
    ''')
    
    # Create Overtime & Transport Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS overtime_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT,
            ot_date TEXT,
            start_time TEXT,
            end_time TEXT,
            needs_transport INTEGER,
            driver_name TEXT,
            driver_mobile TEXT,
            plate_number TEXT,
            route_type TEXT,
            origin TEXT,
            destination TEXT,
            departure_time TEXT,
            return_time TEXT
        )
    ''')
    
    # Create Meeting Rooms Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS meeting_rooms (
            room_number TEXT PRIMARY KEY,
            room_name TEXT,
            capacity INTEGER,
            location TEXT
        )
    ''')
    
    # Create Bookings Table
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

    # Create Fleet Drivers Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS fleet_drivers (
            driver_name TEXT PRIMARY KEY,
            driver_mobile TEXT,
            plate_number TEXT
        )
    ''')

    # Create Daily Transportation Groups Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS transit_groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_date TEXT,
            driver_name TEXT,
            plate_number TEXT,
            passengers TEXT
        )
    ''')
    
    # Seed default Admin if not exists
    cursor.execute("SELECT * FROM users WHERE username='admin'")
    if not cursor.fetchone():
        cursor.execute("INSERT INTO users VALUES ('admin', 'admin123', 'Admin', 'admin@company.com')")
        
    # Seed some sample rooms if completely empty
    cursor.execute("SELECT COUNT(*) FROM meeting_rooms")
    if cursor.fetchone() == 0:
        cursor.execute("INSERT INTO meeting_rooms VALUES ('101', 'Boardroom', 15, '1st Floor')")
        cursor.execute("INSERT INTO meeting_rooms VALUES ('102', 'Huddle Room Alpha', 6, '2nd Floor')")

    # Seed sample drivers if empty
    cursor.execute("SELECT COUNT(*) FROM fleet_drivers")
    if cursor.fetchone() == 0:
        cursor.execute("INSERT INTO fleet_drivers VALUES ('John Doe', '+628111222333', 'B 1234 ABC')")
        cursor.execute("INSERT INTO fleet_drivers VALUES ('Jane Smith', '+628999888777', 'B 5678 XYZ')")
        
    conn.commit()
    conn.close()

# Start DB Structure
init_db()

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

tabs = ["⏰ Overtime & Transport", "👥 Daily Transit Groups", "📅 Meeting Room Booking", "🛠️ System Administration"]
tab1, tab1_b, tab2, tab3 = st.tabs(tabs)

# --- TAB 1: OVERTIME & TRANSPORT ARRANGEMENTS ---
with tab1:
    st.header("Request Overtime & Logistics Tracking")
    
    conn = get_db_connection()
    holidays_df = pd.read_sql_query("SELECT holiday_date FROM holidays", conn)
    drivers = conn.execute("SELECT * FROM fleet_drivers").fetchall()
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
            st.caption("ℹ️ Automatic Rule Applied: **Sunday/Holiday (07:00 - 15:00)** baseline.")
        else:
            default_start = time(17, 30)
            default_end = time(19, 0)
            st.caption("ℹ️ Automatic Rule Applied: **Weekday/Saturday (17:30 - 19:00)** baseline.")
        
        start_time = st.time_input("OT Start Time", value=default_start)
        end_time = st.time_input("OT End Time", value=default_end)
        needs_transport = st.checkbox("Require Individual Transportation Logistics?")
        
    with col2:
        if needs_transport and drivers:
            driver_options = [d['driver_name'] for d in drivers]
            selected_driver = st.selectbox("Select Available Driver", driver_options)
            
            conn = get_db_connection()
            d_info = conn.execute("SELECT * FROM fleet_drivers WHERE driver_name = ?", (selected_driver,)).fetchone()
            conn.close()
            
            driver_name = selected_driver
            driver_mobile = st.text_input("Driver Mobile Phone Number", value=d_info['driver_mobile'])
            plate_number = st.text_input("Car Plate Registration Number", value=d_info['plate_number'])
            
            route_type = st.selectbox("Route Assignment Context", ["Weekday work", "Sunday work", "Sunday shopping", "Holiday Duty"])
            origin = st.text_input("Origin Address", value="Main Corporate Office")
            destination = st.text_input("Target Destination")
            dep_time = st.time_input("Departure Timeline Estimate", value=end_time)
            ret_time = st.time_input("Return Timeline Estimate", value=time(23, 0))
        else:
            driver_name, driver_mobile, plate_number, route_type, origin, destination, dep_time, ret_time = ["", "", "", "", "", "", "", ""]
            if needs_transport and not drivers:
                st.warning("⚠️ No drivers registered in the system yet. Please configure drivers in System Administration.")

    if st.button("Submit New Overtime Request"):
        conn = get_db_connection()
        conn.execute('''
            INSERT INTO overtime_requests (username, ot_date, start_time, end_time, needs_transport, 
            driver_name, driver_mobile, plate_number, route_type, origin, destination, departure_time, return_time)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (st.session_state.username, date_str, start_time.strftime("%H:%M"), end_time.strftime("%H:%M"),
              1 if needs_transport else 0, driver_name, driver_mobile, plate_number, route_type, origin, destination,
              str(dep_time), str(ret_time)))
        
        user_info = conn.execute("SELECT email_recipients FROM users WHERE username=?", (st.session_state.username,)).fetchone()
        conn.commit()
        conn.close()
        
        st.success("🎉 Overtime log successfully submitted!")
        if user_info and user_info['email_recipients']:
            st.info(f"📧 Notification pushed to predetermined dispatch recipients: **{user_info['email_recipients']}**")

    st.subheader("📋 Overtime Submission History Log")
    conn = get_db_connection()
    if st.session_state.role == "Admin":
        ot_df = pd.read_sql_query("SELECT * FROM overtime_requests", conn)
    else:
        ot_df = pd.read_sql_query("SELECT * FROM overtime_requests WHERE username = ?", conn, params=[st.session_state.username])
    conn.close()
    
    if not ot_df.empty:
        st.dataframe(ot_df, use_container_width=True)
        
        buffer_ot = io.BytesIO()
        with pd.ExcelWriter(buffer_ot, engine='openpyxl') as writer:
            ot_df.to_excel(writer, index=False, sheet_name="Overtime Report")
        excel_data_ot = buffer_ot.getvalue()
        
        st.download_button(
            label="📥 Download Overtime History as Excel (.xlsx)",
            data=excel_data_ot,
            file_name=f"Overtime_Report_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- TAB 1B: DAILY TRANSIT GROUPS DESK ---
with tab1_b:
    st.header("👥 Daily Transportation Grouping & Transfers Desk")
    
    group_col1, group_col2 = st.columns(2)
    
    with group_col1:
        st.subheader("➕ Create Daily Transit Group")
        target_group_date = st.date_input("Select Transit Date", value=date.today(), key="tg_date")
        tg_date_str = target_group_date.strftime("%Y-%m-%d")
        
        conn = get_db_connection()
        drivers_list = conn.execute("SELECT * FROM fleet_drivers").fetchall()
        conn.close()
        
        if drivers_list:
            d_options = [d['driver_name'] for d in drivers_list]
            selected_tg_driver = st.selectbox("Assign Driver", d_options, key="tg_driver")
            
            conn = get_db_connection()
            tg_d_info = conn.execute("SELECT * FROM fleet_drivers WHERE driver_name = ?", (selected_tg_driver,)).fetchone()
            conn.close()
            
            st.text(f"Automated Car Plate: {tg_d_info['plate_number']}")
            passenger_input = st.text_area("Passengers List (Separate names with commas)", placeholder="John, Alice, Bob")
            
            if st.button("Provision Transit Group"):
                if passenger_input.strip():
                    conn = get_db_connection()
                    conn.execute('''
                        INSERT INTO transit_groups (group_date, driver_name, plate_number, passengers)
                        VALUES (?, ?, ?, ?)
                    ''', (tg_date_str, selected_tg_driver, tg_d_info['plate_number'], passenger_input.strip()))
                    conn.commit()
                    conn.close()
                    st.success("🎉 Transportation group created successfully.")
                    st.rerun()
                else:
                    st.error("Please add at least one passenger name.")
        else:
            st.warning("No fleet drivers configured.")

    with group_col2:
        st.subheader("🔄 Edit Passengers / Inter-Car Transfers")
        conn = get_db_connection()
        active_groups = conn.execute("SELECT * FROM transit_groups").fetchall()
        conn.close()
        
        if active_groups:
            group_options = {f"ID {g['id']} | {g['group_date']} | Driver: {g['driver_name']}": g['id'] for g in active_groups}
            selected_group_label = st.selectbox("Select Active Group ID to Modify", list(group_options.keys()))
            selected_group_id = group_options[selected_group_label]
            
            conn = get_db_connection()
            selected_group = conn.execute("SELECT * FROM transit_groups WHERE id = ?", (selected_group_id,)).fetchone()
            drivers_edit_list = conn.execute("SELECT * FROM fleet_drivers").fetchall()
            conn.close()
            
            if selected_group:
                edit_passengers = st.text_area("Modify Passenger List (Comma separated)", value=selected_group['passengers'])
                
                st.markdown("**Transfer Group Assignment to Another Driver/Car:**")
                edit_d_options = [d['driver_name'] for d in drivers_edit_list]
                try:
                    current_d_idx = edit_d_options.index(selected_group['driver_name'])
                except:
                    current_d_idx = 0
                    
                transfer_driver = st.selectbox("Transfer to Driver", edit_d_options, index=current_d_idx)
                
                conn = get_db_connection()
                tr_d_info = conn.execute("SELECT * FROM fleet_drivers WHERE driver_name = ?", (transfer_driver,)).fetchone()
                conn.close()
                
                if st.button("Save Transit Group Revisions"):
                    conn = get_db_connection()
                    conn.execute('''
                        UPDATE transit_groups 
                        SET passengers = ?, driver_name = ?, plate_number = ?
                        WHERE id = ?
                    ''', (edit_passengers.strip(), transfer_driver, tr_d_info['plate_number'], selected_group_id))
                    conn.commit()
                    conn.close()
                    st.success("🎉 Group passenger allocations updated seamlessly.")
                    st.rerun()
        else:
            st.info("No transportation groups have been created yet.")

    st.subheader("📊 Active Daily Transit Matrix Log")
    conn = get_db_connection()
    transit_df = pd.read_sql_query("SELECT * FROM transit_groups", conn)
    conn.close()
    if not transit_df.empty:
        st.dataframe(transit_df, use_container_width=True)

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
            recurrence_end = st.date_input("Recurrence End Horizon Target (Max 6 Months Limit)", value=book_date + timedelta(days=7))
            
            if recurrence_end > max_rec_end:
                st.error("⚠️ Rule Restriction Failure: System parameters block automated room recurrence from exceeding a maximum limit of 6 months.")
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
                    st.error(f"❌ Schedule Collision Error! Room {selected_room_num} is already reserved on {t_date_str} during those hours.")
                    conflict_detected = True
                    break
            
            if not conflict_detected:
                for t_date in target_dates:
                    conn.execute('''
                        INSERT INTO room_bookings (room_number, booked_by, booking_date, start_time, end_time, is_recurring, recurrence_end_date)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (selected_room_num, st.session_state.username, t_date.strftime("%Y-%m-%d"), 
                          b_start.strftime("%H:%M"), b_end.strftime("%H:%M"), recurrence, recurrence_end.strftime("%Y-%m-%d")))
                
                user_info = conn.execute("SELECT email_recipients FROM users WHERE username=?", (st.session_state.username,)).fetchone()
                conn.commit()
                st.success(f"🎉 Room assignment established successfully across {len(target_dates)} calendar intervals!")
                if user_info and user_info['email_recipients']:
                    st.info(f"📧 Notification logs dispatched to: **{user_info['email_recipients']}**")
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
        
        buffer_bk = io.BytesIO()
        with pd.ExcelWriter(buffer_bk, engine='openpyxl') as writer:
            bookings_df.to_excel(writer, index=False, sheet_name="Schedules Report")
        excel_data_bk = buffer_bk.getvalue()
        
        st.download_button(
            label="📥 Download Calendar Agenda as Excel (.xlsx)",
            data=excel_data_bk,
            file_name=f"Meeting_Room_Schedules_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- TAB 3: SYSTEM MASTER ADMINISTRATION CONTROL BOARDS ---
with tab3:
    if st.session_state.role != "Admin":
        st.error("🛡️ Restricted Access Control: You lack administrative clearing profiles to view these configuration matrices.")
    else:
        st.header("Admin Control Dashboard Engine")
        
        st.markdown("---")
        st.subheader("🗃️ Master Data Tables CRUD Explorer & Live Grid Editor (Admin Only)")
        table_options = ["users", "holidays", "overtime_requests", "meeting_rooms", "room_bookings", "fleet_drivers", "transit_groups"]
        selected_table = st.selectbox("Choose Database Table to Manage Natively", table_options)
        
        conn = get_db_connection()
        table_df = pd.read_sql_query(f"SELECT * FROM {selected_table}", conn)
        conn.close()
        
        st.markdown(f"👉 **Double-click cells to Edit. Click the '+' icon at the bottom of the grid to Add new rows.**")
        
        # Upgraded to st.data_editor to easily handle add, update, append natively in UI
        edited_df = st.data_editor(table_df, num_rows="dynamic", use_container_width=True, key=f"editor_{selected_table}")
        
        if st.button("💾 Save Grid Changes to Database", key=f"save_{selected_table}"):
            conn = get_db_connection()
            cursor = conn.cursor()
            
            # Wipe the target table completely and rewrite with the edited dataset frame to sync changes cleanly
            cursor.execute(f"DELETE FROM {selected_table}")
            
            # Reinsert rows matching columns dynamically
            columns = edited_df.columns.tolist()
            placeholders = ", ".join(["?"] * len(columns))
            query = f"INSERT INTO {selected_table} ({', '.join(columns)}) VALUES ({placeholders})"
            
            for index, row in edited_df.iterrows():
                row_values = [None if pd.isna(val) else val for val in row.values]
                cursor.execute(query, row_values)
                
            conn.commit()
            conn.close()
            st.success(f"🎉 Live modifications for `{selected_table}` successfully synchronized into the database!")
            st.rerun()
            
        st.markdown("---")
        
        col_adm1, col_adm2 = st.columns(2)
        with col_adm1:
            st.markdown(f"**Alternative Row Purge from `{selected_table}`**")
            id_column_name = "username" if selected_table in ["users", "fleet_drivers"] else ("holiday_date" if selected_table == "holidays" else "room_number" if selected_table == "meeting_rooms" else "id")
            target_row_key = st.text_input(f"Enter row lookup value to drop (Provide {id_column_name}):")
            if st.button(f"Delete Row Row"):
                if target_row_key:
                    conn = get_db_connection()
                    conn.execute(f"DELETE FROM {selected_table} WHERE {id_column_name} = ?", (target_row_key,))
                    conn.commit()
                    conn.close()
                    st.success(f"Row dropped.")
                    st.rerun()
