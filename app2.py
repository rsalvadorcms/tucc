import streamlit as st
import pandas as pd
import sqlite3
import io
import urllib.parse
import os
import re
import shutil
from datetime import datetime, date, timedelta, time

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

# Optional import for PDF rendering
try:
    from reportlab.lib.pagesizes import A3, A4, portrait
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

# Page configuration
st.set_page_config(page_title="Office Operations Portal", layout="wide")

LOGO1_PATH = "logo.png"
LOGO2_PATH = "logo2.png"
BACKUP_DIR = "backups"
NEWS_DIR = "site_news_uploads"
DB_FILE = "office_operations.db"

os.makedirs(BACKUP_DIR, exist_ok=True)
os.makedirs(NEWS_DIR, exist_ok=True)

# ==============================================================================
# 🗄️ DATABASE INITIALIZATION
# ==============================================================================
def get_db_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute('CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password TEXT NOT NULL, role TEXT NOT NULL, email_recipients TEXT, emp_name TEXT)')
    cursor.execute('CREATE TABLE IF NOT EXISTS holidays (holiday_date TEXT PRIMARY KEY, description TEXT)')
    cursor.execute('CREATE TABLE IF NOT EXISTS trips (trip TEXT PRIMARY KEY, trip_name TEXT NOT NULL)')
    cursor.execute('CREATE TABLE IF NOT EXISTS overtime_requests (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, emp_name TEXT, ot_date TEXT, start_time TEXT, end_time TEXT, needs_transport TEXT DEFAULT "Yes", origin TEXT, destination TEXT, departure_time TEXT, return_time TEXT)')
    cursor.execute('CREATE TABLE IF NOT EXISTS meeting_rooms (room_number TEXT PRIMARY KEY, room_name TEXT, capacity INTEGER, location TEXT)')
    cursor.execute('CREATE TABLE IF NOT EXISTS room_bookings (id INTEGER PRIMARY KEY AUTOINCREMENT, room_number TEXT, booked_by TEXT, booking_date TEXT, start_time TEXT, end_time TEXT, is_recurring TEXT, recurrence_end_date TEXT)')
    cursor.execute('CREATE TABLE IF NOT EXISTS cars (car_name TEXT PRIMARY KEY, plate_number TEXT UNIQUE NOT NULL, vehicle TEXT, color TEXT DEFAULT "Black")')
    cursor.execute('CREATE TABLE IF NOT EXISTS fleet_drivers (driver_name TEXT PRIMARY KEY, driver_mobile TEXT)')
    cursor.execute('CREATE TABLE IF NOT EXISTS transit_groups (id INTEGER PRIMARY KEY AUTOINCREMENT, group_name TEXT UNIQUE NOT NULL, driver_name TEXT, etd_1 TEXT, etd_2 TEXT)')
    cursor.execute('CREATE TABLE IF NOT EXISTS transit_passengers (id INTEGER PRIMARY KEY AUTOINCREMENT, group_name TEXT NOT NULL, passengers TEXT NOT NULL)')
    cursor.execute('CREATE TABLE IF NOT EXISTS daily_transit (id INTEGER PRIMARY KEY AUTOINCREMENT, transit_date_start TEXT NOT NULL, transit_date_end TEXT NOT NULL, group_name TEXT DEFAULT "TBA", requested_by TEXT, etd_1 TEXT, etd_2 TEXT, location_from TEXT, location_to TEXT, daily TEXT DEFAULT "No", trip TEXT DEFAULT "Trip A")')
    cursor.execute('CREATE TABLE IF NOT EXISTS site_news (id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL, filename TEXT NOT NULL, file_path TEXT NOT NULL, file_type TEXT NOT NULL, uploaded_by TEXT, upload_date TEXT)')
    
    # Default admin user
    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO users VALUES ('admin', 'admin123', 'Admin', 'admin@company.com', 'System Administrator')")

    # Default trips
    cursor.execute("SELECT COUNT(*) FROM trips")
    if cursor.fetchone()[0] == 0:
        for t_code, t_desc in [("Trip A", "Yard to Yard"), ("Trip B", "Sunday Panbil - Wasco - Panbil"), ("Trip C", "Panbil - Destination - Panbil"), ("Trip D", "Overtime Dispatch Route")]:
            cursor.execute("INSERT INTO trips (trip, trip_name) VALUES (?, ?)", (t_code, t_desc))

    # Default TBA car
    cursor.execute("SELECT COUNT(*) FROM cars WHERE car_name = 'TBA'")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO cars (car_name, plate_number, vehicle, color) VALUES ('TBA', 'TBA', 'TBA', 'TBA')")

    conn.commit()
    conn.close()

init_db()

# ==============================================================================
# 💬 DIALOG HELPER
# ==============================================================================
@st.dialog("Transaction Status")
def show_dialog(title, msg, status="success"):
    if status == "success":
        st.success(f"### {title}")
    else:
        st.error(f"### {title}")
    st.write(msg)
    if st.button("OK", type="primary", use_container_width=True):
        st.session_state.pop('dialog_title', None)
        st.rerun()

if 'dialog_title' in st.session_state:
    show_dialog(st.session_state.dialog_title, st.session_state.get('dialog_msg', ''), st.session_state.get('dialog_type', 'success'))

# ==============================================================================
# 🔐 AUTHENTICATION
# ==============================================================================
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""
    st.session_state.emp_name = ""

if not st.session_state.logged_in:
    st.title("🏢 Office Operations Portal")
    st.subheader("Login to access system")
    with st.form("login_form"):
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        if st.form_submit_button("Log In"):
            conn = get_db_connection()
            user = conn.execute("SELECT * FROM users WHERE username=? AND password=?", (u, p)).fetchone()
            conn.close()
            if user:
                st.session_state.logged_in = True
                st.session_state.username = user['username']
                st.session_state.role = user['role']
                st.session_state.emp_name = user['emp_name'] or user['username']
                st.rerun()
            else:
                st.error("Invalid credentials.")
    st.stop()

# ==============================================================================
# 🗂️ SIDEBAR NAVIGATION & MOBILE STYLING
# ==============================================================================
st.sidebar.title(f"👋 Welcome, {st.session_state.username}")
st.sidebar.info(f"Role: **{st.session_state.role}**")
st.sidebar.markdown("---")

# Strict left alignment + Mobile responsiveness CSS
st.sidebar.markdown(
    """
    <style>
    [data-testid="stSidebar"] button {
        display: flex !important;
        justify-content: flex-start !important;
        text-align: left !important;
        width: 100% !important;
        padding: 10px !important;
    }
    [data-testid="stSidebar"] button p, 
    [data-testid="stSidebar"] button div,
    [data-testid="stSidebar"] button span {
        text-align: left !important;
        justify-content: flex-start !important;
        width: 100% !important;
    }
    @media (max-width: 768px) {
        h1 { font-size: 1.4rem !important; }
        h2 { font-size: 1.2rem !important; }
        h3 { font-size: 1.05rem !important; }
    }
    </style>
    """,
    unsafe_allow_html=True
)

nav_options = [
    "🏠 Daily Transportation Arrangement",
    "📅 Shuttle Timetable",
    "⏰ Overtime & Transport",
    "👥 Transit Groups & Passengers",
    "📅 Daily Transit Dispatch Setup",
    "🏢 Meeting Rooms",
    "📢 Site News",
    "🛠️ System Administration"
]

if 'nav_selection' not in st.session_state:
    st.session_state.nav_selection = nav_options[0]

for opt in nav_options:
    is_active = (st.session_state.nav_selection == opt)
    if is_active:
        st.sidebar.markdown(
            f'<div style="background-color: #FFF2CC; border-left: 5px solid #1F4E78; padding: 10px 12px; margin-bottom: 6px; border-radius: 4px;">{opt}</div>',
            unsafe_allow_html=True
        )
    else:
        if st.sidebar.button(opt, key=f"nav_{opt}", use_container_width=True):
            st.session_state.nav_selection = opt
            st.rerun()

st.sidebar.markdown("---")
if st.sidebar.button("Logout"):
    st.session_state.logged_in = False
    st.rerun()

nav = st.session_state.nav_selection

# ==============================================================================
# 🚀 VIEW ROUTER
# ==============================================================================
if nav == "🏠 Daily Transportation Arrangement":
    st.header("Daily Transportation Arrangement - Passenger List")
    conn = get_db_connection()
    df = pd.read_sql_query('''
        SELECT tg.group_name AS "Car Group", c.vehicle AS "Vehicle Model", c.plate_number AS "Plate Number", 
               c.color AS "Color", tg.driver_name AS "Driver Name", fd.driver_mobile AS "Contact Number", 
               tg.etd_1 AS "ETD 1", tg.etd_2 AS "ETD 2", tp.passengers AS "Passenger Name"
        FROM transit_groups tg
        LEFT JOIN cars c ON tg.group_name = c.car_name
        LEFT JOIN fleet_drivers fd ON tg.driver_name = fd.driver_name
        LEFT JOIN transit_passengers tp ON tg.group_name = tp.group_name
    ''', conn)
    conn.close()
    if not df.empty:
        st.dataframe(df, use_container_width=True)
    else:
        st.info("No transit groups or passenger assignments found.")

elif nav == "📅 Shuttle Timetable":
    st.header("JGC Shuttle Timetable")
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM daily_transit WHERE trip IN ('Trip A', 'Trip B', 'Trip C')", conn)
    conn.close()
    if not df.empty:
        st.dataframe(df, use_container_width=True)
    else:
        st.info("No shuttle schedules configured.")

elif nav == "⏰ Overtime & Transport":
    st.header("Overtime & Transport Request")
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM overtime_requests", conn)
    conn.close()
    st.dataframe(df, use_container_width=True)

elif nav == "👥 Transit Groups & Passengers":
    st.header("Transit Groups & Passengers Management")
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM transit_groups", conn)
    conn.close()
    st.dataframe(df, use_container_width=True)

elif nav == "📅 Daily Transit Dispatch Setup":
    st.header("Daily Transit Dispatch Setup")
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM daily_transit", conn)
    conn.close()
    st.dataframe(df, use_container_width=True)

elif nav == "🏢 Meeting Rooms":
    st.header("Meeting Space Reservations")
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM meeting_rooms", conn)
    conn.close()
    st.dataframe(df, use_container_width=True)

elif nav == "📢 Site News":
    st.header("Site News & Announcements")
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM site_news", conn)
    conn.close()
    st.dataframe(df, use_container_width=True)

elif nav == "🛠️ System Administration":
    st.header("System Administration & Data Management")
    if st.session_state.role not in ["Admin", "Owner"]:
        st.error("Admin access required.")
    else:
        conn = get_db_connection()
        tbl = st.selectbox("Select Table to Edit", ["users", "cars", "fleet_drivers", "holidays", "trips", "meeting_rooms"])
        df = pd.read_sql_query(f"SELECT * FROM {tbl}", conn)
        conn.close()
        
        edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)
        if st.button("Save Changes"):
            conn = get_db_connection()
            conn.execute(f"DELETE FROM {tbl}")
            edited.to_sql(tbl, conn, if_exists="append", index=False)
            conn.commit()
            conn.close()
            st.session_state.dialog_title = "Success"
            st.session_state.dialog_msg = "Table updated successfully."
            st.rerun()