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
        # Opens WhatsApp share menu if no specific phone number is provided
        return f"https://api.whatsapp.com/send?text={encoded_text}"

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
            email_recipients TEXT
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
        CREATE TABLE IF NOT EXISTS fleet_drivers (
            driver_name TEXT PRIMARY KEY,
            driver_mobile TEXT,
            plate_number TEXT
        )
    ''')

    # Transit groups without group_date
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS transit_groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_name TEXT,
            plate_number TEXT,
            passengers TEXT,
            etd_1 TEXT,
            etd_2 TEXT
        )
    ''')

    # Junction Table: daily_transit
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS daily_transit (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            transit_date TEXT NOT NULL,
            transit_group_id INTEGER NOT NULL,
            FOREIGN KEY (transit_group_id) REFERENCES transit_groups (id) ON DELETE CASCADE
        )
    ''')
    
    cursor.execute("SELECT * FROM users WHERE username='admin'")
    if not cursor.fetchone():
        cursor.execute("INSERT INTO users VALUES ('admin', 'admin123', 'Admin', 'admin@company.com')")
        
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

# Schema migration for transition from old schema to daily_transit
def run_migrations():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("PRAGMA table_info(transit_groups)")
    columns = [col[1] for col in cursor.fetchall()]
    
    if "group_date" in columns:
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS transit_groups_new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                driver_name TEXT,
                plate_number TEXT,
                passengers TEXT,
                etd_1 TEXT,
                etd_2 TEXT
            )
        ''')
        cursor.execute('''
            INSERT INTO transit_groups_new (id, driver_name, plate_number, passengers, etd_1, etd_2)
            SELECT id, driver_name, plate_number, passengers, etd_1, etd_2 FROM transit_groups
        ''')
        cursor.execute('''
            INSERT INTO daily_transit (transit_date, transit_group_id)
            SELECT group_date, id FROM transit_groups WHERE group_date IS NOT NULL
        ''')
        cursor.execute("DROP TABLE transit_groups")
        cursor.execute("ALTER TABLE transit_groups_new RENAME TO transit_groups")
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

tabs = ["⏰ Overtime & Transport", "👥 Daily Transit Groups", "📅 Meeting Room Booking", "🛠️ System Administration"]
tab1, tab1_b, tab2, tab3 = st.tabs(tabs)

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
        
        # WhatsApp Share Option for Driver
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
            
            create_etd_1 = st.text_input("ETD 1", value="05:45")
            create_etd_2 = st.text_input("ETD 2", value="17:30")
            passenger_input = st.text_area("Passengers List (Separate names with commas)", placeholder="John, Alice, Bob")
            
            if st.button("Provision Transit Group"):
                if passenger_input.strip():
                    conn = get_db_connection()
                    cursor = conn.cursor()
                    cursor.execute('''
                        INSERT INTO transit_groups (driver_name, plate_number, passengers, etd_1, etd_2)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (selected_tg_driver, tg_d_info['plate_number'], passenger_input.strip(), create_etd_1, create_etd_2))
                    
                    new_group_id = cursor.lastrowid
                    
                    cursor.execute('''
                        INSERT INTO daily_transit (transit_date, transit_group_id)
                        VALUES (?, ?)
                    ''', (tg_date_str, new_group_id))
                    
                    conn.commit()
                    conn.close()
                    st.success("🎉 Transportation group created successfully.")
                    
                    # Generate WhatsApp Share Buttons
                    wa_transit_text = (
                        f"🚌 *DAILY TRANSIT SCHEDULE ({tg_date_str})*\n"
                        f"👤 *Driver:* {selected_tg_driver} ({tg_d_info['plate_number']})\n"
                        f"⏰ *ETD 1:* {create_etd_1} | *ETD 2:* {create_etd_2}\n"
                        f"👥 *Passengers:* {passenger_input.strip()}"
                    )
                    
                    # Direct link to driver's WhatsApp
                    wa_driver_link = generate_whatsapp_link(tg_d_info['driver_mobile'], wa_transit_text)
                    st.link_button("📲 Send Schedule Direct to Driver via WhatsApp", wa_driver_link)
                    
                    # General link to share to any WhatsApp Group
                    wa_group_link = generate_whatsapp_link("", wa_transit_text)
                    st.link_button("📢 Share Schedule to WhatsApp Group", wa_group_link)
                    
                else:
                    st.error("Please add at least one passenger name.")
        else:
            st.warning("No fleet drivers configured.")

    with group_col2:
        st.subheader("🔄 Edit Passengers / Inter-Car Transfers")
        conn = get_db_connection()
        active_groups = conn.execute('''
            SELECT dt.id AS daily_id, dt.transit_date, tg.id AS group_id, tg.driver_name, tg.passengers, tg.etd_1, tg.etd_2
            FROM daily_transit dt
            JOIN transit_groups tg ON dt.transit_group_id = tg.id
        ''').fetchall()
        conn.close()
        
        if active_groups:
            group_options = {f"Schedule #{g['daily_id']} | Date: {g['transit_date']} | Group #{g['group_id']} | Driver: {g['driver_name']}": g for g in active_groups}
            selected_group_label = st.selectbox("Select Active Schedule to Modify", list(group_options.keys()))
            selected_item = group_options[selected_group_label]
            
            conn = get_db_connection()
            drivers_edit_list = conn.execute("SELECT * FROM fleet_drivers").fetchall()
            conn.close()
            
            edit_date = st.date_input("Modify Transit Date", value=datetime.strptime(selected_item['transit_date'], "%Y-%m-%d").date())
            edit_passengers = st.text_area("Modify Passenger List (Comma separated)", value=selected_item['passengers'])
            edit_etd_1 = st.text_input("Modify ETD 1", value=selected_item['etd_1'])
            edit_etd_2 = st.text_input("Modify ETD 2", value=selected_item['etd_2'])
            
            edit_d_options = [d['driver_name'] for d in drivers_edit_list]
            try:
                current_d_idx = edit_d_options.index(selected_item['driver_name'])
            except ValueError:
                current_d_idx = 0
                
            transfer_driver = st.selectbox("Transfer to Driver", edit_d_options, index=current_d_idx)
            
            conn = get_db_connection()
            tr_d_info = conn.execute("SELECT * FROM fleet_drivers WHERE driver_name = ?", (transfer_driver,)).fetchone()
            conn.close()
            
            if st.button("Save Transit Group Revisions"):
                conn = get_db_connection()
                conn.execute('''
                    UPDATE transit_groups 
                    SET passengers = ?, driver_name = ?, plate_number = ?, etd_1 = ?, etd_2 = ?
                    WHERE id = ?
                ''', (edit_passengers.strip(), transfer_driver, tr_d_info['plate_number'], edit_etd_1, edit_etd_2, selected_item['group_id']))
                
                conn.execute('''
                    UPDATE daily_transit
                    SET transit_date = ?
                    WHERE id = ?
                ''', (edit_date.strftime("%Y-%m-%d"), selected_item['daily_id']))
                
                conn.commit()
                conn.close()
                st.success("🎉 Group passenger allocations and schedules updated seamlessly.")
                st.rerun()
        else:
            st.info("No transportation groups have been scheduled yet.")

    st.markdown("---")
    st.subheader("📊 Filter & Export Daily Transit Groups per Date")
    
    filter_export_date = st.date_input("Select Date to Export to Excel", value=date.today(), key="filter_export_date")
    filter_date_str = filter_export_date.strftime("%Y-%m-%d")
    
    conn = get_db_connection()
    filtered_transit_df = pd.read_sql_query('''
        SELECT dt.id AS daily_transit_id, dt.transit_date, tg.id AS group_id, tg.driver_name, tg.plate_number, tg.passengers, tg.etd_1, tg.etd_2
        FROM daily_transit dt
        JOIN transit_groups tg ON dt.transit_group_id = tg.id
        WHERE dt.transit_date = ?
    ''', conn, params=[filter_date_str])
    
    all_transit_df = pd.read_sql_query('''
        SELECT dt.id AS daily_transit_id, dt.transit_date, tg.id AS group_id, tg.driver_name, tg.plate_number, tg.passengers, tg.etd_1, tg.etd_2
        FROM daily_transit dt
        JOIN transit_groups tg ON dt.transit_group_id = tg.id
    ''', conn)
    conn.close()
    
    if not filtered_transit_df.empty:
        st.markdown(f"Records found for **{filter_date_str}**:")
        st.dataframe(filtered_transit_df, use_container_width=True)
        
        # WhatsApp Share button for the daily summary table
        summary_text = f"📋 *TRANSIT SUMMARY FOR {filter_date_str}*\n\n"
        for idx, row in filtered_transit_df.iterrows():
            summary_text += f"🚘 *Car:* {row['driver_name']} ({row['plate_number']})\n⏰ ETD 1: {row['etd_1']} | ETD 2: {row['etd_2']}\n👥 {row['passengers']}\n---\n"
            
        wa_summary_link = generate_whatsapp_link("", summary_text)
        
        col_dl, col_wa = st.columns(2)
        with col_dl:
            buffer_tg = io.BytesIO()
            with pd.ExcelWriter(buffer_tg, engine='openpyxl') as writer:
                filtered_transit_df.to_excel(writer, index=False, sheet_name=f"Transit {filter_date_str}")
            excel_data_tg = buffer_tg.getvalue()
            
            st.download_button(
                label=f"📥 Download {filter_date_str} Transit Groups as Excel (.xlsx)",
                data=excel_data_tg,
                file_name=f"Transit_Groups_{filter_date_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_dl_date_excel"
            )
        with col_wa:
            st.link_button("📢 Share Entire Day's Schedule to WhatsApp Group", wa_summary_link)
    else:
        st.info(f"No custom transportation groups scheduled for {filter_date_str} yet.")
        
    st.markdown("---")
    st.markdown("**All Recorded Groups Log Matrix:**")
    if not all_transit_df.empty:
        st.dataframe(all_transit_df, use_container_width=True)

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
        st.subheader("🗃️ Master Data Tables Inline CRUD Editor (Admin Only)")
        table_options = ["users", "holidays", "overtime_requests", "meeting_rooms", "room_bookings", "fleet_drivers", "transit_groups", "daily_transit"]
        selected_table = st.selectbox("Choose Database Table to Manage", table_options)
        
        conn = get_db_connection()
        table_df = pd.read_sql_query(f"SELECT * FROM {selected_table}", conn)
        conn.close()
        
        st.markdown("💡 *Double-click cells to Edit. Click '+' at the bottom of the grid to Add new rows.*")
        
        edited_df = st.data_editor(table_df, num_rows="dynamic", use_container_width=True, key=f"editor_{selected_table}")
        
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
            st.success(f"🎉 Changes synchronized with '{selected_table}' successfully!")
            st.rerun()
