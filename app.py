import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time

# Page Configuration
st.set_page_config(page_title="Office Management Hub", layout="wide")

# Establish Streamlit Database Connection (Uses secrets under [connections.postgresql])
try:
    # Explicitly pass the connection URL keyword argument directly from secrets
    conn = st.connection("postgresql", type="sql", url=st.secrets["SUPABASE_URL"])
except Exception as e:
    st.error("Database connection configuration missing or invalid. Please check your Streamlit Secrets.")
    st.stop()

# Initialize Database Schema
def init_db():
    with conn.session as session:
        # 1. Users Table
        session.execute("""
            CREATE TABLE IF NOT EXISTS users (
                username TEXT PRIMARY KEY,
                password TEXT NOT NULL,
                role TEXT NOT NULL,
                email_recipients TEXT
            );
        """)
        # Seed default admin if not exists
        res = session.execute("SELECT COUNT(*) FROM users WHERE username = 'admin';").fetchone()
        if res[0] == 0:
            session.execute(
                "INSERT INTO users (username, password, role, email_recipients) VALUES (:u, :p, :r, :e);",
                {"u": "admin", "p": "admin123", "r": "Admin", "e": "admin@company.com"}
            )
        
        # 2. Holidays Table
        session.execute("""
            CREATE TABLE IF NOT EXISTS holidays (
                holiday_date DATE PRIMARY KEY,
                description TEXT NOT NULL
            );
        """)
        
        # 3. Overtime Table (incorporating drivers and routes)
        session.execute("""
            CREATE TABLE IF NOT EXISTS overtime_requests (
                id SERIAL PRIMARY KEY,
                username TEXT NOT NULL,
                ot_date DATE NOT NULL,
                start_time TIME NOT NULL,
                end_time TIME NOT NULL,
                driver_name TEXT,
                driver_mobile TEXT,
                car_plate TEXT,
                route_type TEXT,
                origin TEXT,
                destination TEXT,
                time_of_departure TIME,
                time_of_return TIME
            );
        """)
        
        # 4. Meeting Rooms Table
        session.execute("""
            CREATE TABLE IF NOT EXISTS meeting_rooms (
                room_number TEXT PRIMARY KEY,
                room_name TEXT NOT NULL,
                capacity INTEGER NOT NULL,
                location TEXT NOT NULL
            );
        """)
        # Seed default meeting rooms if table is empty
        res = session.execute("SELECT COUNT(*) FROM meeting_rooms;").fetchone()
        if res[0] == 0:
            session.execute("INSERT INTO meeting_rooms (room_number, room_name, capacity, location) VALUES ('101', 'Boardroom', 12, '1st Floor');")
            session.execute("INSERT INTO meeting_rooms (room_number, room_name, capacity, location) VALUES ('102', 'Huddle Room', 4, '2nd Floor');")

        # 5. Room Bookings Table
        session.execute("""
            CREATE TABLE IF NOT EXISTS room_bookings (
                id SERIAL PRIMARY KEY,
                room_number TEXT NOT NULL,
                booking_date DATE NOT NULL,
                start_time TIME NOT NULL,
                end_time TIME NOT NULL,
                booked_by TEXT NOT NULL,
                purpose TEXT
            );
        """)
        session.commit()

init_db()

# --- Authentication Logic ---
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""

def login_user(username, password):
    user = conn.query("SELECT * FROM users WHERE username = :u AND password = :p;", params={"u": username, "p": password}, ttl=0)
    if not user.empty:
        st.session_state.logged_in = True
        st.session_state.username = username
        st.session_state.role = user.iloc[0]['role']
        return True
    return False

def logout_user():
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""
    st.rerun()

# --- Login UI ---
if not st.session_state.logged_in:
    st.title("🔒 Office Hub Login")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.password_input("Password")
        submit = st.form_submit_button("Login")
        if submit:
            if login_user(username, password):
                st.success(f"Welcome back, {username}!")
                st.rerun()
            else:
                st.error("Invalid Username or Password")
    st.stop()

# --- App Navigation Header ---
st.title("🏢 Office Operations Management Portal")
st.sidebar.markdown(f"**User:** {st.session_state.username} ({st.session_state.role})")
if st.sidebar.button("Logout"):
    logout_user()

tabs = ["📅 Meeting Room Booking", "⏳ Overtime & Transport", "⚙️ Admin Dashboard"]
tab1, tab2, tab3 = st.tabs(tabs)

# ==========================================
# TAB 1: MEETING ROOM BOOKING
# ==========================================
with tab1:
    st.header("Meeting Room Reservation System")
    
    # 1. Fetch available rooms
    rooms_df = conn.query("SELECT * FROM meeting_rooms;", ttl=0)
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("New Booking")
        if rooms_df.empty:
            st.warning("No meeting rooms available. Admin must add a room first.")
        else:
            room_options = {f"{row['room_number']} - {row['room_name']} (Cap: {row['capacity']})": row['room_number'] for _, row in rooms_df.iterrows()}
            selected_room_lbl = st.selectbox("Select Room", list(room_options.keys()))
            selected_room_no = room_options[selected_room_lbl]
            
            booking_date = st.date_input("Meeting Date", datetime.today())
            start_time = st.time_input("Start Time", time(9, 0))
            end_time = st.time_input("End Time", time(10, 0))
            purpose = st.text_input("Meeting Purpose/Subject")
            
            # Recurrence Configuration
            is_recurring = st.checkbox("Is this a recurring meeting?")
            recurrence_type = "None"
            if is_recurring:
                recurrence_type = st.selectbox("Recurrence Interval", ["Daily", "Weekly", "Monthly"])
                st.info("⚠️ Note: Maximum recurrence limit is 6 months down the line.")
                
            if st.button("Check & Book Room"):
                if start_time >= end_time:
                    st.error("Error: End time must occur after the start time.")
                else:
                    # Generate all dates to evaluate based on recurrence rules
                    dates_to_book = [booking_date]
                    max_date = booking_date + timedelta(days=180) # 6 months cap
                    
                    if is_recurring:
                        current_date = booking_date
                        while True:
                            if recurrence_type == "Daily":
                                current_date += timedelta(days=1)
                            elif recurrence_type == "Weekly":
                                current_date += timedelta(weeks=1)
                            elif recurrence_type == "Monthly":
                                # Simple monthly shift estimation
                                current_date += timedelta(days=30)
                                
                            if current_date > max_date:
                                break
                            dates_to_book.append(current_date)
                    
                    # Conflict Check Algorithm
                    has_conflict = False
                    conflicting_date = None
                    
                    for d in dates_to_book:
                        # Simple non-overlapping check syntax compatible with Postgres
                        conflict_query = """
                            SELECT COUNT(*) FROM room_bookings 
                            WHERE room_number = :r 
                            AND booking_date = :d 
                            AND NOT (:end_t <= start_time OR :start_t >= end_time);
                        """
                        res = conn.query(conflict_query, params={"r": selected_room_no, "d": d, "start_t": start_time, "end_t": end_time}, ttl=0)
                        if res.iloc[0, 0] > 0:
                            has_conflict = True
                            conflicting_date = d
                            break
                            
                    if has_conflict:
                        st.error(f"❌ Booking conflict detected on date: {conflicting_date}. Please modify scheduling.")
                    else:
                        # Perform safe transaction commit
                        with conn.session as session:
                            for d in dates_to_book:
                                session.execute("""
                                    INSERT INTO room_bookings (room_number, booking_date, start_time, end_time, booked_by, purpose)
                                    VALUES (:r, :d, :st, :et, :user, :p);
                                """, {"r": selected_room_no, "d": d, "st": start_time, "et": end_time, "user": st.session_state.username, "p": purpose})
                            session.commit()
                        
                        # Fetch email recipients designated for this user profile to trigger notification
                        user_info = conn.query("SELECT email_recipients FROM users WHERE username = :u;", params={"u": st.session_state.username}, ttl=0)
                        recipients = user_info.iloc[0]['email_recipients'] if not user_info.empty else "N/A"
                        
                        st.success(f"🎉 Successfully reserved room for {len(dates_to_book)} session(s)!")
                        st.info(f"📧 Notification sent to recipients configured on your user profile: {recipients}")
                        st.rerun()

    with col2:
        st.subheader("Current Bookings Dashboard")
        all_bookings = conn.query("""
            SELECT b.id, b.room_number, r.room_name, b.booking_date, b.start_time, b.end_time, b.booked_by, b.purpose 
            FROM room_bookings b
            JOIN meeting_rooms r ON b.room_number = r.room_number
            ORDER BY b.booking_date ASC, b.start_time ASC;
        """, ttl=0)
        
        if all_bookings.empty:
            st.info("No active room reservations found.")
        else:
            st.dataframe(all_bookings, use_container_width=True)
            
            # Allow users to delete their own bookings, or Admins to delete any
            st.markdown("**Cancel a Reservation**")
            cancel_id = st.number_input("Enter Booking ID to remove", min_value=1, step=1)
            if st.button("Delete Booking"):
                # Verify permission
                check_owner = conn.query("SELECT booked_by FROM room_bookings WHERE id = :id;", params={"id": cancel_id}, ttl=0)
                if check_owner.empty:
                    st.error("Booking ID not found.")
                elif st.session_state.role == "Admin" or check_owner.iloc[0]['booked_by'] == st.session_state.username:
                    with conn.session as session:
                        session.execute("DELETE FROM room_bookings WHERE id = :id;", {"id": cancel_id})
                        session.commit()
                    st.success(f"Booking #{cancel_id} removed.")
                    st.rerun()
                else:
                    st.error("Permission denied. You can only delete your own bookings.")

# ==========================================
# TAB 2: OVERTIME & TRANSPORT
# ==========================================
with tab2:
    st.header("Overtime Logging & Transportation Coordination")
    
    col_ot1, col_ot2 = st.columns([1, 2])
    
    with col_ot1:
        st.subheader("Log Overtime Session")
        ot_date = st.date_input("Date of Overtime", datetime.today())
        
        # Calculate dynamic start time baseline based on Day of Week and Holiday Table
        is_holiday = not conn.query("SELECT 1 FROM holidays WHERE holiday_date = :d;", params={"d": ot_date}, ttl=0).empty
        is_sunday = (ot_date.weekday() == 6)
        
        if is_sunday or is_holiday:
            default_start = time(7, 0)
            st.caption("ℹ️ Sunday/Holiday detected. Default OT start rules snap to **07:00**.")
        else:
            default_start = time(17, 30)
            st.caption("ℹ️ Standard Workday detected (Mon-Sat). Default OT start rules snap to **17:30**.")
            
        start_t_ot = st.time_input("OT Start Time", default_start)
        end_t_ot = st.time_input("OT End Time", time(21, 0))
        
        st.markdown("---")
        st.subheader("Fleet Coordination Add-on")
        need_driver = st.checkbox("Requires Fleet Assignment / Transportation?")
        
        # Fleet input defaults
        driver_name, driver_mobile, car_plate, route_type, origin, destination = "", "", "", "Weekday work", "", ""
        dep_time, ret_time = time(17, 30), time(21, 0)
        
        if need_driver:
            driver_name = st.text_input("Driver Full Name")
            driver_mobile = st.text_input("Driver Mobile Line")
            car_plate = st.text_input("Vehicle Plate Number")
            route_type = st.selectbox("Route Category / Rule", ["Weekday work", "Sunday work", "Sunday shopping", "Holiday dispatch"])
            origin = st.text_input("Pickup Origin Point")
            destination = st.text_input("Dropoff Destination Point")
            dep_time = st.time_input("Departure Log Estimation", start_t_ot)
            ret_time = st.time_input("Return Log Estimation", end_t_ot)
            
        if st.button("Submit Overtime Request"):
            if start_t_ot >= end_t_ot:
                st.error("Error: Overtime timeline layout is inverted. Fix start/end settings.")
            else:
                with conn.session as session:
                    session.execute("""
                        INSERT INTO overtime_requests (
                            username, ot_date, start_time, end_time, 
                            driver_name, driver_mobile, car_plate, route_type, 
                            origin, destination, time_of_departure, time_of_return
                        ) VALUES (:u, :d, :st, :et, :dn, :dm, :cp, :rt, :ori, :dest, :dep, :ret);
                    """, {
                        "u": st.session_state.username, "d": ot_date, "st": start_t_ot, "et": end_t_ot,
                        "dn": driver_name if need_driver else None,
                        "dm": driver_mobile if need_driver else None,
                        "cp": car_plate if need_driver else None,
                        "rt": route_type if need_driver else None,
                        "ori": origin if need_driver else None,
                        "dest": destination if need_driver else None,
                        "dep": dep_time if need_driver else None,
                        "ret": ret_time if need_driver else None
                    })
                    session.commit()
                
                # Dynamic Email Mock Dispatch Notification Lookup
                user_info = conn.query("SELECT email_recipients FROM users WHERE username = :u;", params={"u": st.session_state.username}, ttl=0)
                recipients = user_info.iloc[0]['email_recipients'] if not user_info.empty else "N/A"
                
                st.success("🎉 Overtime request and transportation schedule successfully logged into central ledger!")
                st.info(f"📧 Notification digest sent automatically to specified profile contacts: {recipients}")
                st.rerun()

    with col_ot2:
        st.subheader("Overtime Ledger & Logs Ledger")
        ot_logs = conn.query("SELECT * FROM overtime_requests ORDER BY ot_date DESC, start_time DESC;", ttl=0)
        if ot_logs.empty:
            st.info("No overtime logging transactions populated yet.")
        else:
            st.dataframe(ot_logs, use_container_width=True)
            
            st.markdown("**Modify or Cancel Log Row Entry**")
            delete_ot_id = st.number_input("Enter Overtime ID to cancel/remove", min_value=1, step=1)
            if st.button("Delete Overtime Request"):
                check_ot_owner = conn.query("SELECT username FROM overtime_requests WHERE id = :id;", params={"id": delete_ot_id}, ttl=0)
                if check_ot_owner.empty:
                    st.error("Overtime record ID not found.")
                elif st.session_state.role == "Admin" or check_ot_owner.iloc[0]['username'] == st.session_state.username:
                    with conn.session as session:
                        session.execute("DELETE FROM overtime_requests WHERE id = :id;", {"id": delete_ot_id})
                        session.commit()
                    st.success(f"Overtime record entry #{delete_ot_id} removed.")
                    st.rerun()
                else:
                    st.error("Permission denied. You can only remove your own log records.")

# ==========================================
# TAB 3: ADMIN DASHBOARD
# ==========================================
with tab3:
    st.header("Administrative Command Dashboard")
    if st.session_state.role != "Admin":
        st.warning("🛑 Access restricted. Administrative clearance role flag missing.")
    else:
        admin_subtab1, admin_subtab2, admin_subtab3 = st.tabs(["👤 Manage Users", "🏖️ Holiday Settings", "🚪 Manage Meeting Rooms"])
        
        # User Administration Subtab
        with admin_subtab1:
            st.subheader("Register System User Profiles")
            with st.form("add_user_form"):
                new_username = st.text_input("New Username Account String")
                new_password = st.text_input("New Password Account String", type="password")
                new_role = st.selectbox("Assign Privilege Target", ["User", "Admin"])
                new_emails = st.text_input("Notification Email Recipients (separated by commas)")
                user_submit = st.form_submit_button("Create User Account Profile")
                
                if user_submit:
                    if not new_username or not new_password:
                        st.error("Username and Password are mandatory fields.")
                    else:
                        try:
                            with conn.session as session:
                                session.execute("""
                                    INSERT INTO users (username, password, role, email_recipients)
                                    VALUES (:u, :p, :r, :e);
                                """, {"u": new_username, "p": new_password, "r": new_role, "e": new_emails})
                                session.commit()
                            st.success(f"Account Profile '{new_username}' added to platform database.")
                        except Exception as ex:
                            st.error(f"Failed to create profile. Username may already exist. Error details: {ex}")
            
            st.markdown("---")
            st.subheader("System Access Directory Profile List")
            system_users = conn.query("SELECT username, role, email_recipients FROM users;", ttl=0)
            st.dataframe(system_users, use_container_width=True)
            
            st.markdown("**Revoke System Access Profile**")
            delete_user_target = st.text_input("Enter EXACT Username string to remove from ledger")
            if st.button("Delete User Profile Row"):
                if delete_user_target == "admin":
                    st.error("Fatal Protection Rule: Seed admin profile line cannot be dropped.")
                else:
                    with conn.session as session:
                        session.execute("DELETE FROM users WHERE username = :u;", {"u": delete_user_target})
                        session.commit()
                    st.success(f"Access revoked for account target '{delete_user_target}'.")
                    st.rerun()

        # Holiday Calendar Configuration Subtab
        with admin_subtab2:
            st.subheader("Configure Calendar Statutory Holidays")
            with st.form("holiday_form"):
                h_date = st.date_input("Holiday Calendar Target", datetime.today())
                h_desc = st.text_input("Statutory Description (e.g. Christmas, New Year)")
                h_submit = st.form_submit_button("Register Holiday Blueprint Date")
                
                if h_submit:
                    try:
                        with conn.session as session:
                            session.execute("INSERT INTO holidays (holiday_date, description) VALUES (:d, :desc);", {"d": h_date, "desc": h_desc})
                            session.commit()
                        st.success(f"Registered calendar holiday rule row for: {h_date}")
                    except Exception:
                        st.error("Holiday rule execution error. Ensure entry rule does not conflict or duplicate baseline logs.")
                        
            st.markdown("---")
            st.subheader("Configured Holiday Calendar Ledgers")
            holidays_df = conn.query("SELECT * FROM holidays ORDER BY holiday_date ASC;", ttl=0)
            if holidays_df.empty:
                st.info("No administrative statutory holiday rules customized yet.")
            else:
                st.dataframe(holidays_df, use_container_width=True)
                
                st.markdown("**Remove Registered Holiday Rule Line**")
                del_h_date = st.date_input("Select Date Rule To Clear", datetime.today(), key="del_h_date")
                if st.button("Delete Holiday Calendar Line"):
                    with conn.session as session:
                        session.execute("DELETE FROM holidays WHERE holiday_date = :d;", {"d": del_h_date})
                        session.commit()
                    st.success(f"Holiday calendar parameter rule removed for: {del_h_date}")
                    st.rerun()

        # Meeting Room Modification Subtab
        with admin_subtab3:
            st.subheader("Add a New Meeting Room")
            with st.form("room_form"):
                r_num = st.text_input("Room Number (Unique ID)")
                r_name = st.text_input("Room Name")
                r_cap = st.number_input("Seating Capacity", min_value=1, step=1, value=10)
                r_loc = st.text_input("Physical Location (e.g., Building A, Floor 3)")
                room_submit = st.form_submit_button("Register Meeting Room")
                
                if room_submit:
                    if not r_num or not r_name:
                        st.error("Room Number and Name are required.")
                    else:
                        try:
                            with conn.session as session:
                                session.execute("""
                                    INSERT INTO meeting_rooms (room_number, room_name, capacity, location)
                                    VALUES (:rn, :name, :cap, :loc);
                                """, {"rn": r_num, "name": r_name, "cap": r_cap, "loc": r_loc})
                                session.commit()
                            st.success(f"Meeting room {r_num} ('{r_name}') registered successfully.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error registering room. ID may already exist: {e}")
                            
            st.markdown("---")
            st.subheader("Manage Current Meeting Rooms")
            st.dataframe(rooms_df, use_container_width=True)
            
            del_room_num = st.text_input("Enter Room Number to Delete")
            if st.button("Delete Meeting Room"):
                with conn.session as session:
                    # Also clean up bookings for that room
                    session.execute("DELETE FROM room_bookings WHERE room_number = :rn;", {"rn": del_room_num})
                    session.execute("DELETE FROM meeting_rooms WHERE room_number = :rn;", {"rn": del_room_num})
                    session.commit()
                st.success(f"Room {del_room_num} and its corresponding bookings have been removed.")
                st.rerun()
