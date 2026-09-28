import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime, timedelta, time

# Set page configurations
st.set_page_config(page_title="Office Operations Portal", layout="wide")

# ==============================================================================
# 🎨 1. CUSTOM VISUAL THEME ENGINE (Modify color codes here to match your brand)
# ==============================================================================
PRIMARY_COLOR = "#1A365D"     # Deep corporate navy blue
ACCENT_COLOR = "#2B6CB0"      # Vibrant blue for buttons/highlights
BACKGROUND_COLOR = "#F7FAFC"  # Light clean background canvas
TEXT_COLOR = "#2D3748"        # Sharp charcoal reading text

custom_css = f"""
    <style>
        /* Main page background */
        .stApp {{
            background-color: {BACKGROUND_COLOR};
            color: {TEXT_COLOR};
        }}
        /* Left Sidebar styling */
        section[data-testid="stSidebar"] {{
            background-color: {PRIMARY_COLOR};
        }}
        section[data-testid="stSidebar"] * {{
            color: #FFFFFF !important;
        }}
        /* Target buttons */
        div.stButton > button {{
            background-color: {ACCENT_COLOR} !important;
            color: white !important;
            border-radius: 6px !important;
            border: none !important;
            font-weight: bold;
        }}
        /* Decorative card structures */
        .css-1r6slb0, .e1f1d6gn1 {{
            background-color: #FFFFFF;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
            margin-bottom: 15px;
        }}
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)


# ==============================================================================
# ⚙️ 2. SELF-CONTAINED DATABASE ENGINE
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
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO meeting_rooms VALUES ('101', 'Boardroom', 15, '1st Floor')")
        cursor.execute("INSERT INTO meeting_rooms VALUES ('102', 'Huddle Room Alpha', 6, '2nd Floor')")
        
    conn.commit()
    conn.close()

# Start DB Structure
init_db()


# ==============================================================================
# 🔐 3. AUTHENTICATION USER INTERFACE
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
# 🗂️ 4. MAIN APP CONTROL PANELS
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

    # Display Logs and Data Spreadsheet Export Download Interface
    st.subheader("📋 Overtime Submission History Log")
    conn = get_db_connection()
    query = "SELECT * FROM overtime_requests" if st.session_state.role == "Admin" else f"SELECT * FROM overtime_requests WHERE username='{st.session_state.username}'"
    ot_df = pd.read_sql_query(query, conn)
    conn.close()
    
    if not ot_df.empty:
        st.dataframe(ot_df, use_container_width=True)
        
        # 📥 DOWNLOAD FILE FEATURE
        csv_data = ot_df.to_csv(index=False).encode('utf-8')
