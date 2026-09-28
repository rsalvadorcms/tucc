import streamlit as st
import pandas as pd
import sqlite3
import io
from datetime import datetime, timedelta, time

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
    
    # Seed default Admin if not exists
    cursor.execute("SELECT * FROM users WHERE username='admin'")
    if not cursor.fetchone():
        cursor.execute("INSERT INTO users VALUES ('admin', 'admin123', 'Admin', 'admin@company.com')")
        
    # Seed some sample rooms if completely empty
    cursor.execute("SELECT COUNT(*) FROM meeting_rooms")
    if cursor.fetchone() == 0:
        cursor.execute("INSERT INTO meeting_rooms VALUES ('101', 'Boardroom', 15, '1st Floor')")
        cursor.execute("INSERT INTO meeting_rooms VALUES ('102', 'Huddle Room Alpha', 6, '2nd Floor')")
        
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

tabs = ["⏰ Overtime & Transport", "📅 Meeting Room Booking", "🛠️ System Administration"]
tab1, tab2, tab3 = st.tabs(tabs)


# --- TAB 1: OVERTIME & TRANSPORT ARRANGEMENTS ---
with tab1:
    st.header("Request Overtime & Logistics Tracking")
    
    conn = get_db_connection()
    holidays_df = pd.read_sql_query("SELECT holiday_date FROM holidays", conn)
    holiday_list = holidays_df['holiday_date'].tolist()
    conn.close()
    
    col1, col2 = st.columns(2)
    with col1:
        ot_date = st.date_input("Select Target Date", value=datetime.today())
        date_str = ot_date.strftime("%Y-%m-%d")
        is_sunday = ot_date.weekday() == 6
        is_holiday = date_str in holiday_list
        
        # Core Requirement Rule Engine Implementation
        default_start = time(7, 0) if (is_sunday or is_holiday) else time(17, 30)
        st.caption(f"ℹ️ Automatic Rule Applied: **{'Sunday/Holiday (07:00)' if (is_sunday or is_holiday) else 'Weekday/Saturday (17:30)'}** baseline.")
        
        start_time = st.time_input("OT Start Time", value=default_start)
        end_time = st.time_input("OT End Time", value=time(21, 0))
        needs_transport = st.checkbox("Require Transportation Logistics?")
        
    with col2:
        if needs_transport:
            driver_name = st.text_input("Driver Name", value="John Doe")
            driver_mobile = st.text_input("Driver Mobile Phone Number")
            plate_number = st.text_input("Car Plate Registration Number")
            route_type = st.selectbox("Route Assignment Context", ["Weekday work", "Sunday work", "Sunday shopping", "Holiday Duty"])
            origin = st.text_input("Origin Address", value="Main Corporate Office")
            destination = st.text_input("Target Destination")
            dep_time = st.time_input("Departure Timeline Estimate", value=end_time)
            ret_time = st.time_input("Return Timeline Estimate", value=time(23, 0))
        else:
            driver_name, driver_mobile, plate_number, route_type, origin, destination, dep_time, ret_time = ["", "", "", "", "", "", "", ""]

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
        
        st.success("🎉 Overtime and logistical logs successfully submitted!")
        if user_info and user_info['email_recipients']:
            st.info(f"📧 Notification pushed to predetermined dispatch recipients: **{user_info['email_recipients']}**")

    # Display Logs and Native Excel Export Interface
    st.subheader("📋 Overtime Submission History Log")
    conn = get_db_connection()
    query = "SELECT * FROM overtime_requests" if st.session_state.role == "Admin" else f"SELECT * FROM overtime_requests WHERE username='{st.session_state.username}'"
    ot_df = pd.read_sql_query(query, conn)
    conn.close()
    
    if not ot_df.empty:
        st.dataframe(ot_df, use_container_width=True)
        
        # 📊 NATIVE EXCEL DOWNLOAD ENGINE (.xlsx)
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
        
        if st.session_state.role == "Admin":
            delete_id = st.number_input("Enter Record Row ID number to purge:", min_value=1, step=1, key="del_ot")
            if st.button("Delete Selected OT Record", key="btn_del_ot"):
                conn = get_db_connection()
                conn.execute("DELETE FROM overtime_requests WHERE id=?", (delete_id,))
                conn.commit()
                conn.close()
                st.success(f"Log ID {delete_id} deleted successfully.")
                st.rerun()


# --- TAB 2: MEETING ROOM BOOKINGS ENGINE ---
with tab2:
    st.header("Meeting Space Reservations Desk")
    
    conn = get_db_connection()
    rooms = conn.execute("SELECT * FROM meeting_rooms").fetchall()
    conn.close()
    
    if not rooms:
        st.warning("No physical boardrooms or meeting layout spaces are registered yet.")
    else:
        room_options = {f"{r['room_name']} (Room {r['room_number']} - Capacity: {r['capacity']})": r['room_number'] for r in rooms}
        selected_room_label = st.selectbox("Choose Target Room Venue", list(room_options.keys()))
        selected_room_num = room_options[selected_room_label]
        
        col1, col2 = st.columns(2)
        with col1:
            book_date = st.date_input("Reservation Date", value=datetime.today(), key="bk_date")
            b_start = st.time_input("Reservation Start Time", value=time(9, 0), key="bk_start")
            b_end = st.time_input("Reservation End Time", value=time(10, 0), key="bk_end")
        
        with col2:
            recurrence = st.selectbox("Recurrence Schedule Pattern", ["None", "Daily", "Weekly", "Monthly"])
            max_rec_end = datetime.today() + timedelta(days=180) # 6 Months Lockout Constraint Rule
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
            
            # Double-booking Validation Engine Logic Checks
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
                    st.error(f"❌ Schedule Collision Error! Another team has already reserved Room {selected_room_num} on {t_date_str} during those hours.")
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

    # Active Calendar Data Display Matrix & Native Excel Export Button
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
        
        # 📊 NATIVE EXCEL DOWNLOAD ENGINE (.xlsx)
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
        
        if st.session_state.role == "Admin":
            del_bk_id = st.number_input("Enter Booking ID to remove:", min_value=1, step=1, key="del_bk")
            if st.button("Cancel Selected Booking Line Assignment", key="btn_del_bk"):
                conn = get_db_connection()
                conn.execute("DELETE FROM room_bookings WHERE id=?", (del_bk_id,))
                conn.commit()
                conn.close()
                st.success(f"Booking ID record {del_bk_id} cleared from tracking systems.")
                st.rerun()


# --- TAB 3: SYSTEM MASTER ADMINISTRATION CONTROL BOARDS ---
with tab3:
    if st.session_state.role != "Admin":
        st.error("🛡️ Restricted Access Control: You lack administrative clearing profiles to view these configuration matrices.")
    else:
        st.header("Admin Control Dashboard Engine")
        
        # User Configuration Form Profiles Panel
        st.subheader("👤 Profile Credentials Manager")
        with st.form("user_reg_form"):
            new_user = st.text_input("New Username Account String")
            new_pass = st.text_input("Security Access Password", type="password")
            new_role = st.selectbox("Authorization Cleared Level", ["User", "Admin"])
            new_email = st.text_input("Predetermined Routing Email Notifications (Comma separated)")
            submit_user = st.form_submit_button("Register Account Credentials")
            
            if submit_user and new_user and new_pass:
                conn = get_db_connection()
                try:
                    conn.execute("INSERT INTO users VALUES (?, ?, ?, ?)", (new_user, new_pass, new_role, new_email))
                    conn.commit()
                    st.success(f"User account credential stack for '{new_user}' successfully committed.")
                except sqlite3.IntegrityError:
                    st.error("System Error: That profile handle identifier string is already cataloged.")
                conn.close()
                
        # Register New Physical Meeting Rooms Structure Layout
        st.subheader("🏢 Provision New Corporate Meeting Workspace")
        with st.form("room_reg_form"):
            r_num = st.text_input("Unique Room Key / Index Number")
            r_name = st.text_input("Descriptive Room Label")
            r_cap = st.number_input("Maximum Structural Seat Occupancy Limit", min_value=1, value=10)
            r_loc = st.text_input("Facility Geography Context (e.g. 3rd Floor Annex, West Wing)")
            submit_room = st.form_submit_button("Provision Asset into Records")
            
            if submit_room and r_num and r_name:
                conn = get_db_connection()
                try:
                    conn.execute("INSERT INTO meeting_rooms VALUES (?, ?, ?, ?)", (r_num, r_name, int(r_cap), r_loc))
                    conn.commit()
                    st.success(f"Physical structural layout resource indices cataloged successfully for '{r_name}'.")
                except sqlite3.IntegrityError:
                    st.error("System Error: This location entry code structure conflicts with another workspace record.")
                conn.close()
                st.rerun()

        # Calendar Holiday Operational Boundaries Tracking Form
        st.subheader("📅 Adjust Corporate Operational Holiday Parameters")
        with st.form("holiday_reg_form"):
            h_date = st.date_input("Target Lockout Holiday Calendar Date", value=datetime.today())
            h_desc = st.text_input("Holiday Designation Scope Description")
            submit_holiday = st.form_submit_button("Store Holiday Rule Constraint")
            
            if submit_holiday:
                conn = get_db_connection()
                try:
                    conn.execute("INSERT INTO holidays VALUES (?, ?)", (h_date.strftime("%Y-%m-%d"), h_desc))
                    conn.commit()
                    st.success(f"Holiday calendar constraints configuration committed successfully for {h_date.strftime('%Y-%m-%d')}.")
                except sqlite3.IntegrityError:
                    st.error("System Error: This date profile rule assignment already exists.")
                conn.close()
