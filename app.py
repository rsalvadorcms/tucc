import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time

# --- 1. DATABASE CONNECTION ---
# Establish connection to Supabase/PostgreSQL via Streamlit Secrets
conn = st.connection("postgresql", type="sql")

# --- 2. TABLE INITIALIZATION ---
# Automatically create the required relational database schema if it doesn't exist
def init_db():
    with conn.session as session:
        # Table: Users
        session.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username VARCHAR(50) UNIQUE NOT NULL,
                password VARCHAR(100) NOT NULL,
                role VARCHAR(20) NOT NULL,
                email_recipients TEXT
            );
        """)
        # Pre-seed Admin if table is completely empty
        res = session.execute("SELECT COUNT(*) FROM users;").fetchone()
        if res and res[0] == 0:
            session.execute(
                "INSERT INTO users (username, password, role, email_recipients) VALUES (:u, :p, :r, :e);",
                {"u": "admin", "p": "admin123", "r": "Admin", "e": "admin@company.com"}
            )
        
        # Table: Holidays
        session.execute("""
            CREATE TABLE IF NOT EXISTS holidays (
                holiday_date DATE PRIMARY KEY,
                description VARCHAR(100)
            );
        """)
        
        # Table: Overtime & Fleet Logistics
        session.execute("""
            CREATE TABLE IF NOT EXISTS overtime_requests (
                id SERIAL PRIMARY KEY,
                username VARCHAR(50) NOT NULL,
                ot_date DATE NOT NULL,
                start_time TIME NOT NULL,
                end_time TIME NOT NULL,
                driver_name VARCHAR(100),
                driver_mobile VARCHAR(50),
                car_plate VARCHAR(50),
                route_type VARCHAR(100),
                origin VARCHAR(100),
                destination VARCHAR(100),
                departure_time TIME,
                return_time TIME
            );
        """)
        
        # Table: Meeting Rooms
        session.execute("""
            CREATE TABLE IF NOT EXISTS meeting_rooms (
                room_number VARCHAR(50) PRIMARY KEY,
                room_name VARCHAR(100) NOT NULL,
                capacity INT NOT NULL,
                location VARCHAR(100) NOT NULL
            );
        """)
        
        # Table: Meeting Room Bookings
        session.execute("""
            CREATE TABLE IF NOT EXISTS room_bookings (
                id SERIAL PRIMARY KEY,
                booking_group_id VARCHAR(100),
                room_number VARCHAR(50) NOT NULL,
                username VARCHAR(50) NOT NULL,
                booking_date DATE NOT NULL,
                start_time TIME NOT NULL,
                end_time TIME NOT NULL,
                purpose VARCHAR(200)
            );
        """)
        session.commit()

init_db()

# --- 3. HELPER LOGIC: CHECK HOLIDAYS & DOUBLE BOOKINGS ---
def is_holiday(check_date):
    df = conn.query("SELECT 1 FROM holidays WHERE holiday_date = :d LIMIT 1;", params={"d": check_date}, ttl=0)
    return not df.empty

def check_booking_conflict(room_number, booking_date, start_time, end_time, exclude_id=None):
    query = """
        SELECT * FROM room_bookings 
        WHERE room_number = :room 
          AND booking_date = :b_date 
          AND NOT (end_time <= :s_time OR start_time >= :e_time)
    """
    params = {"room": room_number, "b_date": booking_date, "s_time": start_time, "e_time": end_time}
    if exclude_id:
        query += " AND id != :ex_id"
        params["ex_id"] = exclude_id
    df = conn.query(query, params=params, ttl=0)
    return not df.empty

# --- 4. AUTHENTICATION STATE INTERFACE ---
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'user' not in st.session_state:
    st.session_state['user'] = None
if 'role' not in st.session_state:
    st.session_state['role'] = None

if not st.session_state['logged_in']:
    st.title("🏢 Corporate Office Operations Portal")
    st.subheader("Secure Access Gateway")
    
    with st.form("login_form"):
        username = st.text_input("User ID")
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Sign In")
        
        if submit:
            df = conn.query("SELECT password, role FROM users WHERE username = :u LIMIT 1;", params={"u": username}, ttl=0)
            if not df.empty and df.iloc[0]['password'] == password:
                st.session_state['logged_in'] = True
                st.session_state['user'] = username
                st.session_state['role'] = df.iloc[0]['role']
                st.rerun()
            else:
                st.error("❌ Invalid User ID or Password.")
    st.stop()

# --- 5. APP INTERFACE & NAVIGATION ---
st.set_page_config(layout="wide")
st.sidebar.title(f"Welcome, {st.session_state['user']}!")
st.sidebar.caption(f"Access Privilege: **{st.session_state['role']}**")

menu = st.sidebar.radio("Navigation Menu", [
    "Meeting Room Booking", 
    "Overtime & Transport Request", 
    "Holiday Settings",
    "User Account Admin"
])

if st.sidebar.button("Log Out"):
    st.session_state['logged_in'] = False
    st.session_state['user'] = None
    st.session_state['role'] = None
    st.rerun()

# -------------------------------------------------------------
# FEATURE MODULE 1: MEETING ROOM BOOKING
# -------------------------------------------------------------
if menu == "Meeting Room Booking":
    st.title("📅 Meeting Room Scheduling Desk")
    
    tab1, tab2, tab3 = st.tabs(["Book a Room", "View & Modify Schedules", "Manage Meeting Rooms (Admin)"])
    
    with tab1:
        st.subheader("New Booking Entry")
        rooms_df = conn.query("SELECT * FROM meeting_rooms;", ttl=0)
        
        if rooms_df.empty:
            st.info("⚠️ No meeting rooms are currently registered. Please register rooms via the administration tab first.")
        else:
            room_options = {f"{row['room_number']} - {row['room_name']} (Cap: {row['capacity']})": row['room_number'] for _, row in rooms_df.iterrows()}
            selected_room_label = st.selectbox("Select Room Target", list(room_options.keys()))
            room_num = room_options[selected_room_label]
            
            purpose = st.text_input("Meeting Purpose/Topic")
            start_date = st.date_input("Start Date / First Occurence", datetime.now().date())
            t_start = st.time_input("Start Time", time(9, 0))
            t_end = st.time_input("End Time", time(10, 0))
            
            recurrence = st.selectbox("Recurrence Type", ["None", "Daily", "Weekly", "Monthly"])
            
            if t_start >= t_end:
                st.error("❌ Start time must precede your designated end time.")
            else:
                if st.button("Confirm Booking Schedule"):
                    # Process date schedules based on recurrence constraints (Max 6 Months)
                    dates_to_book = []
                    max_date = start_date + timedelta(days=180)
                    
                    if recurrence == "None":
                        dates_to_book.append(start_date)
                    elif recurrence == "Daily":
                        curr = start_date
                        while curr <= max_date:
                            dates_to_book.append(curr)
                            curr += timedelta(days=1)
                    elif recurrence == "Weekly":
                        curr = start_date
                        while curr <= max_date:
                            dates_to_book.append(curr)
                            curr += timedelta(weeks=1)
                    elif recurrence == "Monthly":
                        curr = start_date
                        while curr <= max_date:
                            dates_to_book.append(curr)
                            # Handle variable month lengths roughly by shifting forward 30 days
                            curr += timedelta(days=30)
                    
                    # Validate all target schedules against database blocks for conflicts
                    conflicts = []
                    for d in dates_to_book:
                        if check_booking_conflict(room_num, d, t_start, t_end):
                            conflicts.append(d.strftime('%Y-%m-%d'))
                    
                    if conflicts:
                        st.error(f"❌ Booking conflict encountered! The room is already occupied on these dates: {', '.join(conflicts[:5])}{'...' if len(conflicts)>5 else ''}")
                    else:
                        import uuid
                        group_id = str(uuid.uuid4())[:8]
                        with conn.session as session:
                            for d in dates_to_book:
                                session.execute("""
                                    INSERT INTO room_bookings (booking_group_id, room_number, username, booking_date, start_time, end_time, purpose)
                                    VALUES (:g, :r, :u, :d, :st, :et, :p);
                                """, {"g": group_id, "r": room_num, "u": st.session_state['user'], "d": d, "st": t_start, "et": t_end, "p": purpose})
                            session.commit()
                        
                        # Generate notification text
                        user_info = conn.query("SELECT email_recipients FROM users WHERE username = :u LIMIT 1;", params={"u": st.session_state['user']}, ttl=0)
                        recipients = user_info.iloc[0]['email_recipients'] if not user_info.empty else "N/A"
                        
                        st.success(f"🎉 Booking successfully recorded! Sequence initialized over {len(dates_to_book)} booking days.")
                        st.info(f"📧 **Automated Notification Summary sent to:** [{recipients}]\n\n**Content:** Room {room_num} booked by {st.session_state['user']} for '{purpose}' on {start_date} at {t_start}-{t_end}.")
    
    with tab2:
        st.subheader("Active Bookings Dashboard")
        filter_room = st.text_input("Filter View by Room Number (Leave empty to view all)")
        
        query = "SELECT * FROM room_bookings"
        params = {}
        if filter_room:
            query += " WHERE room_number = :r"
            params["r"] = filter_room
        query += " ORDER BY booking_date ASC, start_time ASC;"
        
        bookings_df = conn.query(query, params=params, ttl=0)
        
        if bookings_df.empty:
            st.write("No meetings scheduled matching your filter requirements.")
        else:
            st.dataframe(bookings_df, use_container_width=True)
            
            st.markdown("---")
            st.subheader("Modify / Cancel Existing Booking")
            booking_id = st.number_input("Enter Booking ID to update/remove", min_value=1, step=1)
            
            target_booking = conn.query("SELECT * FROM room_bookings WHERE id = :id LIMIT 1;", params={"id": booking_id}, ttl=0)
            if not target_booking.empty:
                st.write(f"Selected Booking: **{target_booking.iloc[0]['purpose']}** on **{target_booking.iloc[0]['booking_date']}**")
                
                col1, col2 = st.columns(2)
                with col1:
                    new_purpose = st.text_input("Modify Purpose", value=target_booking.iloc[0]['purpose'])
                    new_date = st.date_input("Modify Date", value=pd.to_datetime(target_booking.iloc[0]['booking_date']).date())
                    new_start = st.time_input("Modify Start Time", value=pd.to_datetime(str(target_booking.iloc[0]['start_time'])).time())
                    new_end = st.time_input("Modify End Time", value=pd.to_datetime(str(target_booking.iloc[0]['end_time'])).time())
                    
                    if st.button("Apply Changes"):
                        if check_booking_conflict(target_booking.iloc[0]['room_number'], new_date, new_start, new_end, exclude_id=booking_id):
                            st.error("❌ Alteration conflicts with an existing schedule footprint.")
                        else:
                            with conn.session as session:
                                session.execute("""
                                    UPDATE room_bookings 
                                    SET purpose = :p, booking_date = :d, start_time = :st, end_time = :et 
                                    WHERE id = :id;
                                """, {"p": new_purpose, "d": new_date, "st": new_start, "et": new_end, "id": booking_id})
                                session.commit()
                            st.success("🔄 Booking successfully modified.")
                            st.rerun()
                            
                with col2:
                    st.write("Danger Zone Operations")
                    if st.button("🗑️ Delete Single Booking Entry", key="del_single"):
                        with conn.session as session:
                            session.execute("DELETE FROM room_bookings WHERE id = :id;", {"id": booking_id})
                            session.commit()
                        st.success("Entry removed completely from log records.")
                        st.rerun()
                        
                    group_id_val = target_booking.iloc[0]['booking_group_id']
                    if group_id_val and st.button("💥 Cancel Whole Recurring Sequence", key="del_group"):
                        with conn.session as session:
                            session.execute("DELETE FROM room_bookings WHERE booking_group_id = :g;", {"g": group_id_val})
                            session.commit()
                        st.success("Entire related calendar chain cleared.")
                        st.rerun()
            else:
                st.caption("Provide a valid ID from the table above to reveal modification options.")

    with tab3:
        if st.session_state['role'] != "Admin":
            st.error("🔒 Only administrative profiles possess permissions to register new structural meeting rooms.")
        else:
            st.subheader("Register Core Structural Meeting Room")
            with st.form("add_room_form"):
                r_num = st.text_input("Room Number (Unique Identifier, e.g. CONF-401)")
                r_name = st.text_input("Display Room Name")
                r_cap = st.number_input("Capacity Limits", min_value=1, value=10)
                r_loc = st.text_input("Building Location / Floor Level")
                submit_room = st.form_submit_button("Add Asset to Infrastructure")
                
                if submit_room:
                    if not r_num or not r_name:
                        st.error("Please supply valid identification variables.")
                    else:
                        with conn.session as session:
                            session.execute("""
                                INSERT INTO meeting_rooms (room_number, room_name, capacity, location)
                                VALUES (:num, :name, :cap, :loc)
                                ON CONFLICT (room_number) DO UPDATE SET room_name = :name, capacity = :cap, location = :loc;
                            """, {"num": r_num, "name": r_name, "cap": r_cap, "loc": r_loc})
                            session.commit()
                        st.success(f"Room Asset '{r_num}' has been committed to infrastructure records.")
                        st.rerun()

# -------------------------------------------------------------
# FEATURE MODULE 2: OVERTIME & FLEET LOGISTICS
# -------------------------------------------------------------
elif menu == "Overtime & Transport Request":
    st.title("🚗 Overtime Requests & Logistics Deck")
    
    t1, t2 = st.tabs(["Log OT Request", "Review & Track Requests"])
    
    with t1:
        st.subheader("New Overtime & Transportation Booking Log")
        
        ot_date = st.date_input("Target Date for Log", datetime.now().date())
        
        # Calculate dynamic logic rules based on calendar target classification
        is_holiday_flag = is_holiday(ot_date)
        is_sunday = (ot_date.weekday() == 6)
        
        if is_holiday_flag or is_sunday:
            rule_context = "Sunday / Holiday Schedule (07:00 Baseline Rule)"
            default_start = time(7, 0)
        else:
            rule_context = "Standard Weekday Schedule (17:30 Default Rule)"
            default_start = time(17, 30)
            
        st.info(f"📅 **Context Engine Diagnosis:** Verified as **{rule_context}**")
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("##### ⏱️ Hours Logged")
            start_time = st.time_input("OT Start window", default_start)
            end_time = st.time_input("OT Termination window", time(21, 0))
        
        with col2:
            st.markdown("##### 🚙 Fleet Desk Assignment (Optional)")
            driver_name = st.text_input("Driver Name")
            driver_mobile = st.text_input("Driver Contact Line")
            car_plate = st.text_input("Vehicle License Plate")
            route_type = st.selectbox("Route Configuration Profile", [
                "Weekday work", "Sunday work", "Sunday shopping", "Holiday dispatch", "N/A"
            ])
            origin = st.text_input("Origin Point", "Main Office Block")
            destination = st.text_input("Dropoff Destination Address")
            dep_time = st.time_input("Estimated Fleet Departure", start_time)
            ret_time = st.time_input("Estimated Fleet Return", end_time)
            
        if start_time >= end_time:
            st.error("❌ End time window must be placed chronologically after the initialization window.")
        else:
            if st.button("Publish Log Entry"):
                with conn.session as session:
                    session.execute("""
                        INSERT INTO overtime_requests (
                            username, ot_date, start_time, end_time, driver_name, 
                            driver_mobile, car_plate, route_type, origin, destination, departure_time, return_time
                        ) VALUES (:u, :d, :st, :et, :dn, :dm, :cp, :rt, :ori, :dest, :dept, :rett);
                    """, {
                        "u": st.session_state['user'], "d": ot_date, "st": start_time, "et": end_time,
                        "dn": driver_name, "dm": driver_mobile, "cp": car_plate, "rt": route_type,
                        "ori": origin, "dest": destination, "dept": dep_time, "rett": ret_time
                    })
                    session.commit()
                
                # Fetch target emails linked to layout profile structure
                user_info = conn.query("SELECT email_recipients FROM users WHERE username = :u LIMIT 1;", params={"u": st.session_state['user']}, ttl=0)
                recipients = user_info.iloc[0]['email_recipients'] if not user_info.empty else "N/A"
                
                st.success("🎉 Your Overtime Request and Fleet assignment sheet have been published successfully.")
                st.info(f"📧 **Automated Corporate Notification dispatched to:** [{recipients}]\n\n**Content Summary:** OT Request by {st.session_state['user']} logged for {ot_date} starting at {start_time}.")

    with t2:
        st.subheader("Corporate Logs")
        
        # Admins see everything, regular users see only their own items
        if st.session_state['role'] == "Admin":
            ot_df = conn.query("SELECT * FROM overtime_requests ORDER BY ot_date DESC;", ttl=0)
        else:
            ot_df = conn.query("SELECT * FROM overtime_requests WHERE username = :u ORDER BY ot_date DESC;", params={"u": st.session_state['user']}, ttl=0)
            
        if ot_df.empty:
            st.write("No overtime database rows found matching user visibility scopes.")
        else:
            st.dataframe(ot_df, use_container_width=True)
            
            st.markdown("---")
            st.subheader("Manage Database Rows")
            row_to_manage = st.number_input("Specify Entry ID for removal or alteration", min_value=1, step=1, key="ot_row_id")
            
            target_ot = conn.query("SELECT * FROM overtime_requests WHERE id = :id LIMIT 1;", params={"id": row_to_manage}, ttl=0)
            if not target_ot.empty:
                # Enforce rule: regular users cannot modify other profiles
                if st.session_state['role'] != "Admin" and target_ot.iloc[0]['username'] != st.session_state['user']:
                    st.error("🔒 Security Exception: Access Denied. Profile matching constraints broken.")
                else:
                    if st.button("🗑️ Purge Entry Record", key="del_ot"):
                        with conn.session as session:
                            session.execute("DELETE FROM overtime_requests WHERE id = :id;", {"id": row_to_manage})
                            session.commit()
                        st.success("Log item successfully expunged from primary ledgers.")
                        st.rerun()
            else:
                st.caption("Provide an active ID value to reveal modifications tools.")

# -------------------------------------------------------------
# FEATURE MODULE 3: HOLIDAY MANAGEMENT
# -------------------------------------------------------------
elif menu == "Holiday Settings":
    st.title("📅 Corporate Calendar & Holiday Registry")
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("Register Holiday Event")
        if st.session_state['role'] != "Admin":
            st.error("🔒 Administrative clear level required to alter calendar settings.")
        else:
            with st.form("holiday_form"):
                h_date = st.date_input("Holiday Calendar Date Target")
                h_desc = st.text_input("Event Classification (e.g. New Year's Day)")
                submit_h = st.form_submit_button("Commit Holiday to Master Configuration")
                
                if submit_h:
                    with conn.session as session:
                        session.execute("""
                            INSERT INTO holidays (holiday_date, description) VALUES (:d, :desc)
                            ON CONFLICT (holiday_date) DO UPDATE SET description = :desc;
                        """, {"d": h_date, "desc": h_desc})
                        session.commit()
                    st.success("Holiday record saved.")
                    st.rerun()
                    
    with col2:
        st.subheader("Active Calendar Footprints Registered")
        holidays_df = conn.query("SELECT * FROM holidays ORDER BY holiday_date ASC;", ttl=0)
        if holidays_df.empty:
            st.write("No custom company calendar events declared yet.")
        else:
            st.dataframe(holidays_df, use_container_width=True)
            
            if st.session_state['role'] == "Admin":
                st.markdown("##### Remove Calendar Footprint")
                del_h_date = st.date_input("Target Date to Remove", value=holidays_df.iloc[0]['holiday_date'])
                if st.button("Delete Selected Holiday"):
                    with conn.session as session:
                        session.execute("DELETE FROM holidays WHERE holiday_date = :d;", {"d": del_h_date})
                        session.commit()
                    st.success("Holiday baseline calendar cleared successfully.")
                    st.rerun()

# -------------------------------------------------------------
# FEATURE MODULE 4: USER MANAGER CONTROL
# -------------------------------------------------------------
elif menu == "User Account Admin":
    st.title("👥 User Profile & System Credentials Admin")
    
    if st.session_state['role'] != "Admin":
        st.error("🔒 System Administrator Clearance Level mandatory to open this panel.")
    else:
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.subheader("Add / Update User Profile Account")
            with st.form("user_reg_form"):
                new_user = st.text_input("Desired User ID (Plaintext login handle)")
                new_pass = st.text_input("Access Password Token", type="password")
                new_role = st.selectbox("Role Assignment Matrix", ["User", "Admin"])
                new_emails = st.text_area("Default Notification Mailboxes (Comma-separated addresses list)")
                submit_user = st.form_submit_button("Save User Account")
                
                if submit_user:
                    if not new_user or not new_pass:
                        st.error("Username or password credentials cannot remain blank.")
                    else:
                        with conn.session as session:
                            session.execute("""
                                INSERT INTO users (username, password, role, email_recipients)
                                VALUES (:u, :p, :r, :e)
                                ON CONFLICT (username) DO UPDATE SET password = :p, role = :r, email_recipients = :e;
                            """, {"u": new_user, "p": new_pass, "r": new_role, "e": new_emails})
                            session.commit()
                        st.success(f"System profile for user '{new_user}' committed successfully.")
                        st.rerun()
                        
        with col2:
            st.subheader("Registered Active Profile Registry")
            users_df = conn.query("SELECT id, username, role, email_recipients FROM users ORDER BY username ASC;", ttl=0)
            st.dataframe(users_df, use_container_width=True)
            
            st.markdown("---")
            st.subheader("Purge Profile Control")
            user_id_to_del = st.number_input("Target Database Row ID to Delete", min_value=1, step=1)
            
            if st.button("💥 Permanent Deletion of User Account"):
                # Safety feature: Prevent the active user from deleting themselves
                check_self = conn.query("SELECT username FROM users WHERE id = :id LIMIT 1;", params={"id": user_id_to_del}, ttl=0)
                if not check_self.empty and check_self.iloc[0]['username'] == st.session_state['user']:
                    st.error("❌ Safeguard warning: Self-deletion block invoked. You cannot delete your own active session account.")
                else:
                    with conn.session as session:
                        session.execute("DELETE FROM users WHERE id = :id;", {"id": user_id_to_del})
                        session.commit()
                    st.success("Target profile cleared.")
                    st.rerun()
