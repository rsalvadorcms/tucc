import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime, timedelta, time

# Set page config
st.set_page_config(page_title="Office Portal", layout="wide")

# 1. Self-contained Database initialization (Zero Config Required!)
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
        
    # Seed some sample rooms if empty
    cursor.execute("SELECT COUNT(*) FROM meeting_rooms")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO meeting_rooms VALUES ('101', 'Boardroom', 15, '1st Floor')")
        cursor.execute("INSERT INTO meeting_rooms VALUES ('102', 'Huddle Room Alpha', 6, '2nd Floor')")
        
    conn.commit()
    conn.close()

# Initialize DB structure immediately
init_db()

# 2. Authentication Logic
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""

if not st.session_state.logged_in:
    st.title("🏢 Office Operations Portal")
    st.subheader("Login to access system panels")
    
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
                st.error("Invalid username or password.")
    st.stop()

# --- Main App Interface (Logged In) ---
st.sidebar.title(f"👋 Welcome, {st.session_state.username}")
st.sidebar.info(f"Role: {st.session_state.role}")
if st.sidebar.button("Logout"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""
    st.rerun()

tabs = ["⏰ Overtime & Transport", "📅 Meeting Room Booking", "🛠️ System Administration"]
tab1, tab2, tab3 = st.tabs(tabs)

# --- TAB 1: OVERTIME & TRANSPORT ---
with tab1:
    st.header("Request Overtime & Logistics")
    
    # Load Holidays
    conn = get_db_connection()
    holidays_df = pd.read_sql_query("SELECT holiday_date FROM holidays", conn)
    holiday_list = holidays_df['holiday_date'].tolist()
    
    col1, col2 = st.columns(2)
    with col1:
        ot_date = st.date_input("Select Date", value=datetime.today())
        date_str = ot_date.strftime("%Y-%m-%d")
        is_sunday = ot_date.weekday() == 6
        is_holiday = date_str in holiday_list
        
        # Rule implementation: Snap starting baselines
        default_start = time(7, 0) if (is_sunday or is_holiday) else time(17, 30)
        st.caption(f"Rule Target Detected: {'Sunday/Holiday (07:00)' if (is_sunday or is_holiday) else 'Weekday/Saturday (17:30)'}")
        
        start_time = st.time_input("OT Start Time", value=default_start)
        end_time = st.time_input("OT End Time", value=time(21, 0))
        
        needs_transport = st.checkbox("Require Transportation Assignment?")
        
    with col2:
        if needs_transport:
            driver_name = st.text_input("Driver Name", value="John Doe")
            driver_mobile = st.text_input("Driver Mobile Line")
            plate_number = st.text_input("Car Plate Number")
            route_type = st.selectbox("Route Category", ["Weekday work", "Sunday work", "Sunday shopping", "Holiday Dispatch"])
            origin = st.text_input("Origin Point", value="Main Office")
            destination = st.text_input("Destination Address")
            dep_time = st.time_input("Departure Time", value=end_time)
            ret_time = st.time_input("Estimated Return Time", value=time(23, 0))
        else:
            driver_name, driver_mobile, plate_number, route_type, origin, destination, dep_time, ret_time = ["", "", "", "", "", "", "", ""]

    if st.button("Submit Overtime Request"):
        conn = get_db_connection()
        conn.execute('''
            INSERT INTO overtime_requests (username, ot_date, start_time, end_time, needs_transport, 
            driver_name, driver_mobile, plate_number, route_type, origin, destination, departure_time, return_time)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (st.session_state.username, date_str, start_time.strftime("%H:%M"), end_time.strftime("%H:%M"),
              1 if needs_transport else 0, driver_name, driver_mobile, plate_number, route_type, origin, destination,
              str(dep_time), str(ret_time)))
        
        # Email Simulation Trigger
        user_info = conn.execute("SELECT email_recipients FROM users WHERE username=?", (st.session_state.username,)).fetchone()
        conn.commit()
        conn.close()
        
        st.success("🎉 Overtime Request recorded successfully!")
        if user_info and user_info['email_recipients']:
            st.info(f"📧 Notification auto-routed to predetermined recipients: **{user_info['email_recipients']}**")

    # Display / Manage Logs
    st.subheader("Your Overtime History Log")
    conn = get_db_connection()
    query = "SELECT * FROM overtime_requests" if st.session_state.role == "Admin" else f"SELECT * FROM overtime_requests WHERE username='{st.session_state.username}'"
    ot_df = pd.read_sql_query(query, conn)
    conn.close()
    
    if not ot_df.empty:
        st.dataframe(ot_df, use_container_width=True)
        if st.session_state.role == "Admin":
            delete_id = st.number_input("Enter ID row row to purge:", min_value=1, step=1, key="del_ot")
            if st.button("Delete OT Record", key="btn_del_ot"):
                conn = get_db_connection()
                conn.execute("DELETE FROM overtime_requests WHERE id=?", (delete_id,))
                conn.commit()
                conn.close()
                st.success(f"Row {delete_id} deleted successfully.")
                st.rerun()

# --- TAB 2: MEETING ROOM BOOKING ---
with tab2:
    st.header("Meeting Room Desk & Scheduling")
    
    conn = get_db_connection()
    rooms = conn.execute("SELECT * FROM meeting_rooms").fetchall()
    conn.close()
    
    if not rooms:
        st.warning("No meeting rooms configured yet. Admins can register rooms in the Administration panel.")
    else:
        room_options = {f"{r['room_name']} (Room {r['room_number']} - Cap: {r['capacity']})": r['room_number'] for r in rooms}
        selected_room_label = st.selectbox("Choose Target Room", list(room_options.keys()))
        selected_room_num = room_options[selected_room_label]
        
        col1, col2 = st.columns(2)
        with col1:
            book_date = st.date_input("Booking Target Date", value=datetime.today(), key="bk_date")
            b_start = st.time_input("Booking Start Time", value=time(9, 0), key="bk_start")
            b_end = st.time_input("Booking End Time", value=time(10, 0), key="bk_end")
        
        with col2:
            recurrence = st.selectbox("Recurrence Plan Strategy", ["None", "Daily", "Weekly", "Monthly"])
            max_rec_end = datetime.today() + timedelta(days=180) # 6 months restriction
            recurrence_end = st.date_input("Recurrence Stop Baseline (Max 6 Months Limit)", value=book_date + timedelta(days=7))
            
            if recurrence_end > max_rec_end:
                st.error("⚠️ Rule validation failed: Recurrence cannot exceed a maximum safety horizon of 6 months.")
                st.stop()

        if st.button("Confirm Room Booking"):
            # Compute targets dates based on selected strategy
            target_dates = [book_date]
            if recurrence != "None":
                current_date = book_date
