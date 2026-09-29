import streamlit as st
import pandas as pd
import sqlite3
import io
import urllib.parse
from datetime import datetime, date, timedelta, time

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

def get_db_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. users table with emp_name
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

    # 6. fleet_drivers table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS fleet_drivers (
            driver_name TEXT PRIMARY KEY,
            driver_mobile TEXT,
            plate_number TEXT
        )
    ''')

    # 7. transit_groups table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS transit_groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_name TEXT UNIQUE NOT NULL,
            driver_name TEXT,
            etd_1 TEXT,
            etd_2 TEXT,
            FOREIGN KEY (driver_name) REFERENCES fleet_drivers(driver_name) ON DELETE SET NULL
        )
    ''')

    # 8. transit_passengers table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS transit_passengers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_name TEXT NOT NULL,
            passengers TEXT NOT NULL,
            FOREIGN KEY (group_name) REFERENCES transit_groups(group_name) ON DELETE CASCADE
        )
    ''')

    # 9. daily_transit table
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
        cursor.execute("INSERT INTO fleet_drivers VALUES ('John Doe', '+628111222333', 'B 1234 ABC')")
        cursor.execute("INSERT INTO fleet_drivers VALUES ('Jane Smith', '+628999888777', 'B 5678 XYZ')")
        
    conn.commit()
    conn.close()

init_db()

# Schema migration helper
def run_migrations():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Add emp_name column to users if missing
    cursor.execute("PRAGMA table_info(users)")
    user_cols = [col[1] for col in cursor.fetchall()]
    if "emp_name" not in user_cols:
        cursor.execute("ALTER TABLE users ADD COLUMN emp_name TEXT")
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
            st.caption("ℹ️ Baseline Rule: **Sunday/Holiday (07:00 - 15:00)**.")
        else:
            default_start = time(17, 30)
            default_end = time(19, 0)
            st.caption("ℹ️ Baseline Rule: **Weekday/Saturday (17:30 - 19:00)**.")
        
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
        
        if needs_transport and driver_mobile:
            ot_msg = (
                f"🚗 *NEW OVERTIME TRANSPORT REQUEST*\n"
                f"👤 *Passenger:* {st.session_state.username}\n"
                f"📅 *Date:* {date_str}\n"
                f"⏰ *OT Hours:* {start_time.strftime('%H:%M')} - {end_time.strftime('%H:%M')}\n"
                f"📍 *Route:* {origin} ➡️ {destination}\n"
                f"🚘 *Car/Plate:* {plate_number}\n"
                f"🛫 *Est. Departure:* {dep_time}"
            )
            wa_driver_url = generate_whatsapp_link(driver_mobile, ot_msg)
            st.link_button("📲 Notify Driver via WhatsApp", wa_driver_url)

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
    employees_list = [u['emp_name'] for u in conn.execute("SELECT emp_name FROM users WHERE emp_name IS NOT NULL AND emp_name != ''").fetchall()]
    groups_list = [g['group_name'] for g in conn.execute("SELECT group_name FROM transit_groups").fetchall()]
    conn.close()

    col1, col2 = st.columns(2)
    
    # Section A: Define Transit Groups
    with col1:
        st.subheader("1. Create Transit Group")
        g_name = st.text_input("Group Name (e.g., Alpha Shuttles, Route 1)", key="g_name_input")
        selected_driver = st.selectbox("Assign Driver (from Fleet Drivers)", drivers_list if drivers_list else ["No drivers available"])
        etd_1 = st.text_input("ETD 1", value="05:45")
        etd_2 = st.text_input("ETD 2", value="17:30")
        
        if st.button("Save Transit Group"):
            if g_name.strip() and selected_driver != "No drivers available":
                try:
                    conn = get_db_connection()
                    conn.execute("INSERT INTO transit_groups (group_name, driver_name, etd_1, etd_2) VALUES (?, ?, ?, ?)",
                                 (g_name.strip(), selected_driver, etd_1, etd_2))
                    conn.commit()
                    conn.close()
                    st.success(f"Group '{g_name}' created successfully!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("A group with this name already exists.")
            else:
                st.error("Please provide a valid Group Name and select a Driver.")

    # Section B: Assign Passengers to Transit Groups
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
    groups_df = pd.read_sql_query('''
        SELECT tg.id, tg.group_name, tg.driver_name, fd.plate_number, fd.driver_mobile, tg.etd_1, tg.etd_2,
               GROUP_CONCAT(tp.passengers, ', ') AS passengers
        FROM transit_groups tg
        LEFT JOIN fleet_drivers fd ON tg.driver_name = fd.driver_name
        LEFT JOIN transit_passengers tp ON tg.group_name = tp.group_name
        GROUP BY tg.id
    ''', conn)
    conn.close()
    
    if not groups_df.empty:
        st.dataframe(groups_df, use_container_width=True)

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
        st.subheader("📢 Share Schedule")
        filter_date = st.date_input("Filter Schedule Date", value=date.today(), key="filter_sched_date")
        filter_date_str = filter_date.strftime("%Y-%m-%d")
        
        conn = get_db_connection()
        daily_df = pd.read_sql_query('''
            SELECT dt.id AS daily_id, dt.transit_date, tg.group_name, tg.driver_name, fd.plate_number, fd.driver_mobile,
                   tg.etd_1, tg.etd_2, GROUP_CONCAT(tp.passengers, ', ') AS passengers
            FROM daily_transit dt
            JOIN transit_groups tg ON dt.group_name = tg.group_name
            LEFT JOIN fleet_drivers fd ON tg.driver_name = fd.driver_name
            LEFT JOIN transit_passengers tp ON tg.group_name = tp.group_name
            WHERE dt.transit_date = ?
            GROUP BY dt.id
        ''', conn, params=[filter_date_str])
        conn.close()
        
        if not daily_df.empty:
            summary_text = f"📋 *TRANSIT DISPATCH SCHEDULE ({filter_date_str})*\n\n"
            for _, row in daily_df.iterrows():
                summary_text += f"🚌 *Group:* {row['group_name']}\n🚘 *Driver:* {row['driver_name']} ({row['plate_number']})\n⏰ ETD 1: {row['etd_1']} | ETD 2: {row['etd_2']}\n👥 Passengers: {row['passengers']}\n---\n"
                
            wa_link = generate_whatsapp_link("", summary_text)
            st.link_button("📢 Share Schedule to WhatsApp Group", wa_link)

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
        
        st.markdown("---")
        st.subheader("🗃️ Master Data Tables Inline CRUD Editor")
        table_options = ["users", "holidays", "overtime_requests", "meeting_rooms", "room_bookings", "fleet_drivers", "transit_groups", "transit_passengers", "daily_transit"]
        selected_table = st.selectbox("Choose Database Table to Manage", table_options)
        
        conn = get_db_connection()
        table_df = pd.read_sql_query(f"SELECT * FROM {selected_table}", conn)
        
        # Helper lists for dynamic dropdown configuration in st.data_editor
        drivers_list = [d['driver_name'] for d in conn.execute("SELECT driver_name FROM fleet_drivers").fetchall()]
        groups_list = [g['group_name'] for g in conn.execute("SELECT group_name FROM transit_groups").fetchall()]
        emp_list = [u['emp_name'] for u in conn.execute("SELECT emp_name FROM users WHERE emp_name IS NOT NULL AND emp_name != ''").fetchall()]
        conn.close()
        
        # Configure selectbox columns for specific tables
        column_config = {}
        if selected_table == "transit_groups":
            column_config["driver_name"] = st.column_config.SelectboxColumn("Driver Name", options=drivers_list)
        elif selected_table == "transit_passengers":
            column_config["group_name"] = st.column_config.SelectboxColumn("Group Name", options=groups_list)
            column_config["passengers"] = st.column_config.SelectboxColumn("Passenger", options=emp_list)
        elif selected_table == "daily_transit":
            column_config["group_name"] = st.column_config.SelectboxColumn("Group Name", options=groups_list)

        st.markdown("💡 *Edit cells or use dropdowns where configured. Click '+' to add rows.*")
        
        edited_df = st.data_editor(table_df, num_rows="dynamic", use_container_width=True, column_config=column_config, key=f"editor_{selected_table}")
        
        if st.button(f"Save Grid Changes to Database ({selected_table})"):
            conn = get_db_connection()
            cursor = conn.cursor()
            
            cursor.execute(f"DELETE FROM {selected_table}")
            
            for _, row in edited_df.iterrows():
                columns = [k for k in row.keys() if row[k] is not None]
                values = [row[k] for k in columns]
                placeholders = ", ".join(["?"] * len(columns))
                col_names = ", ".join(columns)
                
                cursor.execute(f"INSERT INTO {selected_table} ({col_names}) VALUES ({placeholders})", values)
                
            conn.commit()
            conn.close()
            st.success(f"🎉 Database '{selected_table}' updated successfully!")
            st.rerun()
