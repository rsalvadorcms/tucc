import os
import sqlite3
from datetime import datetime, timedelta
import streamlit as st
import pandas as pd

# --- DATABASE SETUP ---
DB_FILE = "office_hub.db"

def get_db_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        
        # Users Table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL,
                email TEXT NOT NULL,
                notification_recipients TEXT
            )
        ''')
        
        # Drivers & Cars Table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS fleet (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                driver_name TEXT NOT NULL,
                driver_phone TEXT NOT NULL,
                car_plate TEXT NOT NULL,
                status TEXT DEFAULT 'Available'
            )
        ''')
        
        # Holidays Table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS holidays (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                holiday_date TEXT UNIQUE NOT NULL,
                description TEXT
            )
        ''')
        
        # Overtime & Transport Requests Table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS overtime_requests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                date TEXT NOT NULL,
                start_time TEXT NOT NULL,
                end_time TEXT NOT NULL,
                needs_transport INTEGER DEFAULT 0,
                fleet_id INTEGER,
                route_type TEXT,
                origin TEXT,
                destination TEXT,
                departure_time TEXT,
                return_time TEXT,
                FOREIGN KEY(user_id) REFERENCES users(id),
                FOREIGN KEY(fleet_id) REFERENCES fleet(id)
            )
        ''')
        
        # Meeting Rooms Table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS meeting_rooms (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                room_number TEXT UNIQUE NOT NULL,
                room_name TEXT NOT NULL,
                capacity INTEGER NOT NULL,
                location TEXT NOT NULL
            )
        ''''')
        
        # Meeting Bookings Table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS room_bookings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                room_id INTEGER,
                user_id INTEGER,
                booking_date TEXT NOT NULL,
                start_time TEXT NOT NULL,
                end_time TEXT NOT NULL,
                purpose TEXT,
                group_id TEXT,
                FOREIGN KEY(room_id) REFERENCES meeting_rooms(id),
                FOREIGN KEY(user_id) REFERENCES users(id)
            )
        ''')
        
        # Seed default Admin account if empty
        cursor.execute("SELECT COUNT(*) FROM users")
        if cursor.fetchone()[0] == 0:
            cursor.execute("INSERT INTO users (username, password, role, email) VALUES ('admin', 'admin123', 'Admin', 'admin@company.com')")
            
        # Seed default meeting rooms if empty
        cursor.execute("SELECT COUNT(*) FROM meeting_rooms")
        if cursor.fetchone()[0] == 0:
            cursor.execute("INSERT INTO meeting_rooms (room_number, room_name, capacity, location) VALUES ('101', 'Boardroom', 15, '1st Floor')")
            cursor.execute("INSERT INTO meeting_rooms (room_number, room_name, capacity, location) VALUES ('202', 'Alpha Room', 6, '2nd Floor')")
            
        # Seed default fleet if empty
        cursor.execute("SELECT COUNT(*) FROM fleet")
        if cursor.fetchone()[0] == 0:
            cursor.execute("INSERT INTO fleet (driver_name, driver_phone, car_plate) VALUES ('John Doe', '+62812345678', 'B 1234 ABC')")
            cursor.execute("INSERT INTO fleet (driver_name, driver_phone, car_plate) VALUES ('Jane Smith', '+6287654321', 'B 5678 DEF')")
            
        conn.commit()

init_db()

# --- HELPER FUNCTIONS ---
def check_room_conflict(room_id, date, start_time, end_time, exclude_booking_id=None):
    with get_db_connection() as conn:
        query = '''
            SELECT id FROM room_bookings
            WHERE room_id = ? AND booking_date = ? 
            AND NOT (end_time <= ? OR start_time >= ?)
        '''
        params = [room_id, date, start_time, end_time]
        if exclude_booking_id:
            query += " AND id != ?"
            params.append(exclude_booking_id)
        return conn.execute(query, params).fetchone() is not None

def is_holiday_or_sunday(date_str):
    with get_db_connection() as conn:
        holiday = conn.execute("SELECT 1 FROM holidays WHERE holiday_date = ?", (date_str,)).fetchone()
    if holiday:
        return True
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return dt.weekday() == 6 # 6 is Sunday

def simulate_email(subject, body, recipients):
    st.info(f"📧 **Simulated Email Sent to:** `{recipients}`\n\n**Subject:** {subject}\n\n{body}")

# --- APP LAYOUT & AUTHENTICATION ---
st.set_page_config(page_title="Corporate Hub Portal", layout="wide", page_icon="🏢")
st.title("🏢 Corporate Hub Portal")

if 'user' not in st.session_state:
    st.session_state.user = None

if not st.session_state.user:
    st.subheader("Login to your Account")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login")
        if submitted:
            with get_db_connection() as conn:
                user = conn.execute("SELECT * FROM users WHERE username = ? AND password = ?", (username, password)).fetchone()
                if user:
                    st.session_state.user = dict(user)
                    st.rerun()
                else:
                    st.error("Invalid credentials.")
    st.stop()

# --- NAVIGATION SIDEBAR ---
user_info = st.session_state.user
st.sidebar.markdown(f"### Welcome, **{user_info['username']}** (`{user_info['role']}`)")
if st.sidebar.button("Logout"):
    st.session_state.user = None
    st.rerun()

menu = st.sidebar.radio("Navigation Menu", [
    "📅 Meeting Room Booking", 
    "⏳ Overtime & Transport Request", 
    "🚖 Fleet Management",
    "🎉 Holiday Master", 
    "👥 User Management"
])

# ==============================================================================
# 1. MEETING ROOM BOOKING
# ==============================================================================
if menu == "📅 Meeting Room Booking":
    st.header("📅 Meeting Room Management")
    tab1, tab2, tab3 = st.tabs(["Book a Room", "View & Filter Calendars", "Manage Existing Bookings"])
    
    with get_db_connection() as conn:
        rooms = conn.execute("SELECT * FROM meeting_rooms").fetchall()
        room_options = {f"{r['room_name']} (Room {r['room_number']}, Cap: {r['capacity']})": r['id'] for r in rooms}

    with tab1:
        if not room_options:
            st.warning("No meeting rooms available. Add rooms in Admin mode or DB.")
        else:
            st.subheader("Create a Booking")
            with st.form("booking_form"):
                selected_room = st.selectbox("Select Meeting Room", list(room_options.keys()))
                room_id = room_options[selected_room]
                
                b_date = st.date_input("Date", datetime.today())
                start_t = st.time_input("Start Time", value=datetime.strptime("09:00", "%H:%M").time())
                end_t = st.time_input("End Time", value=datetime.strptime("10:00", "%H:%M").time())
                purpose = st.text_area("Purpose of Meeting")
                
                recurrence = st.selectbox("Recurrence Strategy", ["None", "Daily", "Weekly", "Monthly"])
                rec_duration = st.slider("Recurrence Duration (Months)", 1, 6, 1) if recurrence != "None" else 0
                
                submit_booking = st.form_submit_button("Book Room")
                
                if submit_booking:
                    s_t_str = start_t.strftime("%H:%M")
                    e_t_str = end_t.strftime("%H:%M")
                    
                    if s_t_str >= e_t_str:
                        st.error("End time must be after start time.")
                    else:
                        # Build dates sequence
                        dates_to_book = []
                        start_date = b_date
                        if recurrence == "None":
                            dates_to_book.append(start_date.strftime("%Y-%m-%d"))
                        else:
                            end_date = start_date + timedelta(days=30 * rec_duration)
                            current_date = start_date
                            while current_date <= end_date:
                                dates_to_book.append(current_date.strftime("%Y-%m-%d"))
                                if recurrence == "Daily":
                                    current_date += timedelta(days=1)
                                elif recurrence == "Weekly":
                                    current_date += timedelta(weeks=1)
                                elif recurrence == "Monthly":
                                    # Simple approximation for monthly step
                                    current_date += timedelta(days=30)
                        
                        # Validate conflicts for all instances
                        conflicts = []
                        for d in dates_to_book:
                            if check_room_conflict(room_id, d, s_t_str, e_t_str):
                                conflicts.append(d)
                                
                        if conflicts:
                            st.error(f"Booking Conflict detected on these dates: {', '.join(conflicts)}. Please choose another time slot.")
                        else:
                            # Save bookings
                            group_id = f"G-{int(datetime.now().timestamp())}"
                            with get_db_connection() as conn:
                                for d in dates_to_book:
                                    conn.execute('''
                                        INSERT INTO room_bookings (room_id, user_id, booking_date, start_time, end_time, purpose, group_id)
                                        VALUES (?, ?, ?, ?, ?, ?, ?)
                                    ''', (room_id, user_info['id'], d, s_t_str, e_t_str, purpose, group_id))
                                conn.commit()
                            st.success(f"Successfully booked room for {len(dates_to_book)} days!")
                            
                            if user_info['notification_recipients']:
                                simulate_email(
                                    f"New Meeting Room Booking: {selected_room}",
                                    f"Room: {selected_room}\nDates: {', '.join(dates_to_book[:3])}...\nTime: {s_t_str} - {e_t_str}\nBooked by: {user_info['username']}",
                                    user_info['notification_recipients']
                                )
    
    with tab2:
        st.subheader("Filter & View Schedules")
        if room_options:
            filter_room = st.selectbox("Filter by Room", ["All"] + list(room_options.keys()))
            with get_db_connection() as conn:
                q = '''
                    SELECT rb.id, mr.room_name, mr.room_number, rb.booking_date, rb.start_time, rb.end_time, u.username, rb.purpose
                    FROM room_bookings rb
                    JOIN meeting_rooms mr ON rb.room_id = mr.id
                    JOIN users u ON rb.user_id = u.id
                '''
                if filter_room != "All":
                    q += f" WHERE rb.room_id = {room_options[filter_room]}"
                q += " ORDER BY rb.booking_date ASC, rb.start_time ASC"
                df_bookings = pd.read_sql_query(q, conn)
                st.dataframe(df_bookings, use_container_width=True)

    with tab3:
        st.subheader("Modify / Delete Bookings")
        with get_db_connection() as conn:
            user_filter_clause = "" if user_info['role'] == "Admin" else f" WHERE rb.user_id = {user_info['id']}"
            q = f'''
                SELECT rb.id, mr.room_name, rb.booking_date, rb.start_time, rb.end_time, rb.purpose
                FROM room_bookings rb
                JOIN meeting_rooms mr ON rb.room_id = mr.id
                {user_filter_clause}
            '''
            mod_df = pd.read_sql_query(q, conn)
            
        if mod_df.empty:
            st.info("No bookings available for modification.")
        else:
            st.dataframe(mod_df, use_container_width=True)
            b_to_delete = st.selectbox("Select Booking Record ID to Actions", mod_df['id'].tolist())
            
            if st.button("🔴 Delete Selected Booking"):
                with get_db_connection() as conn:
                    conn.execute("DELETE FROM room_bookings WHERE id = ?", (b_to_delete,))
                    conn.commit()
                st.success(f"Booking {b_to_delete} removed successfully.")
                st.rerun()

# ==============================================================================
# 2. OVERTIME & TRANSPORT REQUEST
# ==============================================================================
elif menu == "⏳ Overtime & Transport Request":
    st.header("⏳ Overtime Tracking & Fleet Allocations")
    ot_tab1, ot_tab2 = st.tabs(["Log OT & Request Transport", "History & Archives"])
    
    with ot_tab1:
        st.subheader("File New Overtime Request")
        ot_date = st.date_input("Overtime Date", datetime.today())
        ot_date_str = ot_date.strftime("%Y-%m-%d")
        
        # Calculate policy standard hours automatically
        special_day = is_holiday_or_sunday(ot_date_str)
        default_start_str = "07:00" if special_day else "17:30"
        st.info(f"💡 **Policy Check**: Target date counts as a **{'Sunday/Holiday' if special_day else 'Standard Weekday'}**. Dynamic policy sets OT calculation to start at **{default_start_str}**.")
        
        with st.form("ot_form"):
            t_start = st.time_input("Actual OT Start Time", datetime.strptime(default_start_str, "%H:%M").time())
            t_end = st.time_input("Actual OT End Time", datetime.strptime("21:00", "%H:%M").time())
            
            st.markdown("---")
            needs_transport = st.checkbox("Check here if you require an Office Car & Driver Arrangement")
            
            with get_db_connection() as conn:
                fleet_units = conn.execute("SELECT * FROM fleet").fetchall()
                fleet_options = {f"{f['driver_name']} [{f['car_plate']}]": f['id'] for f in fleet_units}
                
            fleet_choice = st.selectbox("Assign Fleet Unit", list(fleet_options.keys())) if fleet_options else None
            route_type = st.selectbox("Route Profile / Schedule Category", ["Weekday work", "Sunday work", "Sunday shopping", "Holiday Shift"])
            origin = st.text_input("Origin Address", "HQ Office")
            destination = st.text_input("Destination Point")
            dep_time = st.time_input("Fleet Departure Time", value=datetime.strptime("18:00", "%H:%M").time())
            ret_time = st.time_input("Fleet Return Time Estimated", value=datetime.strptime("22:00", "%H:%M").time())
            
            submit_ot = st.form_submit_button("Submit Form & Log Request")
            
            if submit_ot:
                f_id = fleet_options[fleet_choice] if (needs_transport and fleet_choice) else None
                with get_db_connection() as conn:
                    conn.execute('''
                        INSERT INTO overtime_requests (
                            user_id, date, start_time, end_time, needs_transport, fleet_id,
                            route_type, origin, destination, departure_time, return_time
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        user_info['id'], ot_date_str, t_start.strftime("%H:%M"), t_end.strftime("%H:%M"),
                        1 if needs_transport else 0, f_id, route_type, origin, destination,
                        dep_time.strftime("%H:%M") if needs_transport else None,
                        ret_time.strftime("%H:%M") if needs_transport else None
                    ))
                    conn.commit()
                st.success("Overtime log successfully saved!")
                
                if user_info['notification_recipients']:
                    simulate_email(
                        f"New Overtime filed by {user_info['username']}",
                        f"Date: {ot_date_str}\nHours: {t_start.strftime('%H:%M')} - {t_end.strftime('%H:%M')}\nTransport Arranged: {needs_transport}",
                        user_info['notification_recipients']
                    )

    with ot_tab2:
        st.subheader("Stored Log Records")
        with get_db_connection() as conn:
            user_filter_clause = "" if user_info['role'] == "Admin" else f" WHERE otr.user_id = {user_info['id']}"
            q = f'''
                SELECT otr.id, u.username, otr.date, otr.start_time, otr.end_time, otr.needs_transport, 
                       fl.driver_name, fl.car_plate, otr.route_type, otr.origin, otr.destination
                FROM overtime_requests otr
                JOIN users u ON otr.user_id = u.id
                LEFT JOIN fleet fl ON otr.fleet_id = fl.id
                {user_filter_clause}
            '''
            df_ot = pd.read_sql_query(q, conn)
            st.dataframe(df_ot, use_container_width=True)
            
        if not df_ot.empty and st.button("🔴 Purge Selected Log Row"):
            target_id = st.selectbox("Row ID to Remove", df_ot['id'].tolist(), key="del_ot_select")
            with get_db_connection() as conn:
                conn.execute("DELETE FROM overtime_requests WHERE id = ?", (target_id,))
                conn.commit()
            st.success(f"Row {target_id} purged.")
            st.rerun()

# ==============================================================================
# 3. FLEET MANAGEMENT
# ==============================================================================
elif menu == "🚖 Fleet Management":
    st.header("🚖 Fleet Configuration Desk")
    fl_tab1, fl_tab2 = st.tabs(["Active Fleet Registry", "Register Fleet Unit"])
    
    with fl_tab1:
        with get_db_connection() as conn:
            df_fleet = pd.read_sql_query("SELECT * FROM fleet", conn)
        st.dataframe(df_fleet, use_container_width=True)
        
        if user_info['role'] == "Admin" and not df_fleet.empty:
            st.markdown("---")
            st.subheader("Modify Fleet Information")
            f_to_mod = st.selectbox("Select ID to Edit/Delete", df_fleet['id'].tolist())
            row_data = df_fleet[df_fleet['id'] == f_to_mod].iloc[0]
            
            with st.form("edit_fleet"):
                d_name = st.text_input("Driver Full Name", row_data['driver_name'])
                d_phone = st.text_input("Contact Number", row_data['driver_phone'])
                c_plate = st.text_input("Car Plate Number", row_data['car_plate'])
                col1, col2 = st.columns(2)
                with col1:
                    save_f = st.form_submit_button("Update Data Row")
                with col2:
                    del_f = st.form_submit_button("🚨 Wipe Registry Row")
                    
                if save_f:
                    with get_db_connection() as conn:
                        conn.execute("UPDATE fleet SET driver_name=?, driver_phone=?, car_plate=? WHERE id=?", (d_name, d_phone, c_plate, f_to_mod))
                        conn.commit()
                    st.success("Fleet changes stored.")
                    st.rerun()
                if del_f:
                    with get_db_connection() as conn:
                        conn.execute("DELETE FROM fleet WHERE id=?", (f_to_mod,))
                        conn.commit()
                    st.success("Fleet entity purged.")
                    st.rerun()

    with fl_tab2:
        st.subheader("Add Driver-Car Profile")
        with st.form("add_fleet_form"):
            new_d = st.text_input("Driver Full Name")
            new_p = st.text_input("Driver Mobile / Phone Contact")
            new_c = st.text_input("Car Plate Number")
            if st.form_submit_button("Save Unit To Registry"):
                if new_d and new_p and new_c:
                    with get_db_connection() as conn:
                        conn.execute("INSERT INTO fleet (driver_name, driver_phone, car_plate) VALUES (?, ?, ?)", (new_d, new_p, new_c))
                        conn.commit()
                    st.success("New fleet unit deployed to database.")
                    st.rerun()
                else:
                    st.error("Fill all input fields completely.")

# ==============================================================================
# 4. HOLIDAY MASTER
# ==============================================================================
elif menu == "🎉 Holiday Master":
    st.header("🎉 Holiday Configuration Desk")
    col_l, col_r = st.columns([1, 2])
    
    with col_l:
        st.subheader("Add Corporate Holiday")
        with st.form("holiday_form"):
            h_date = st.date_input("Calendar Holiday Date", datetime.today())
            h_desc = st.text_input("Description (e.g. New Year Holiday)")
            if st.form_submit_button("Append Holiday"):
                try:
                    with get_db_connection() as conn:
                        conn.execute("INSERT INTO holidays (holiday_date, description) VALUES (?, ?)", (h_date.strftime("%Y-%m-%d"), h_desc))
                        conn.commit()
                    st.success("Holiday locked into system rules.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("This date already exists in the calendar index.")
                    
    with col_r:
        st.subheader("Indexed Core Holidays")
        with get_db_connection() as conn:
            df_h = pd.read_sql_query("SELECT * FROM holidays ORDER BY holiday_date ASC", conn)
        st.dataframe(df_h, use_container_width=True)
        
        if not df_h.empty and user_info['role'] == 'Admin':
            h_to_del = st.selectbox("Select Target ID to Remove", df_h['id'].tolist())
            if st.button("Remove Selected Holiday"):
                with get_db_connection() as conn:
                    conn.execute("DELETE FROM holidays WHERE id = ?", (h_to_del,))
                    conn.commit()
                st.success("Holiday rule removed.")
                st.rerun()

# ==============================================================================
# 5. USER MANAGEMENT & SETTINGS
# ==============================================================================
elif menu == "👥 User Management":
    st.header("👥 Accounts and Recipient Subscriptions")
    
    if user_info['role'] != 'Admin':
        st.warning("🔒 User Registry and role management is exclusively limited to Corporate Portal Administrators.")
        
        # Self settings for non-admins
        st.subheader("Your Notification Contacts")
        with st.form("self_notify_form"):
            recip = st.text_input("Predefined Target Email Recipients (Comma Separated)", user_info['notification_recipients'] or "")
            if st.form_submit_button("Update My Notifications Settings"):
                with get_db_connection() as conn:
                    conn.execute("UPDATE users SET notification_recipients = ? WHERE id = ?", (recip, user_info['id']))
                    conn.commit()
                st.session_state.user['notification_recipients'] = recip
                st.success("Settings updated successfully.")
    else:
        u_tab1, u_tab2 = st.tabs(["System Accounts View", "Create New User Access Profile"])
        with u_tab1:
            with get_db_connection() as conn:
                df_users = pd.read_sql_query("SELECT id, username, role, email, notification_recipients FROM users", conn)
            st.dataframe(df_users, use_container_width=True)
            
            st.markdown("---")
            st.subheader("Quick Profile Modification")
            user_sel = st.selectbox("Select User Row ID", df_users['id'].tolist())
            matched_user = df_users[df_users['id'] == user_sel].iloc[0]
            
            with st.form("admin_user_mod"):
                u_email = st.text_input("Primary Account Email", matched_user['email'])
                u_role = st.selectbox("Privilege Classification Level", ["User", "Admin"], index=["User", "Admin"].index(matched_user['role']))
                u_recip = st.text_input("Automatic Recipient String List", matched_user['notification_recipients'] or "")
                
                if st.form_submit_button("Commit Accounts Modification"):
                    with get_db_connection() as conn:
                        conn.execute("UPDATE users SET email=?, role=?, notification_recipients=? WHERE id=?", (u_email, u_role, u_recip, user_sel))
                        conn.commit()
                    st.success("User access table updated successfully.")
                    st.rerun()
                    
        with u_tab2:
            st.subheader("Register New Operational / Admin Credentials")
            with st.form("new_user_form"):
                n_user = st.text_input("Desired Username Name Key")
                n_pass = st.text_input("Security Assignment Password", type="password")
                n_role = st.selectbox("Role Rank Class", ["User", "Admin"])
                n_mail = st.text_input("User Corporate Mailbox")
                
                if st.form_submit_button("Create Account Profile"):
                    if n_user and n_pass and n_mail:
                        try:
                            with get_db_connection() as conn:
                                conn.execute("INSERT INTO users (username, password, role, email) VALUES (?, ?, ?, ?)", (n_user, n_pass, n_role, n_mail))
                                conn.commit()
                            st.success(f"Account for user '{n_user}' deployed perfectly.")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("That username string key is already registered inside our corporate cluster.")
                    else:
                        st.error("Provide data inside every input box.")
