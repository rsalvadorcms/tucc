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
    
    # Seed default Admin if not exists
    cursor.execute("SELECT * FROM users WHERE username='admin'")
    if not cursor.fetchone():
        cursor.execute("INSERT INTO users VALUES ('admin', 'admin123', 'Admin', 'admin@company.com')")
        
    # Seed some sample rooms if completely empty
    cursor.execute("SELECT COUNT(*) FROM meeting_rooms")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO meeting_rooms VALUES ('101', 'Boardroom', 15, '1st Floor')")
        cursor.execute("INSERT INTO meeting_rooms VALUES ('102', 'Huddle Room Alpha', 6, '2nd Floor')")

    # Seed some sample drivers if completely empty
    cursor.execute("SELECT COUNT(*) FROM fleet_drivers")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO fleet_drivers VALUES ('John Doe', '+62812345678', 'BP 1234 XY')")
        cursor.execute("INSERT INTO fleet_drivers VALUES ('Jane Smith', '+62876543210', 'BP 5678 AB')")
        
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
    
    # Fetch drivers for dropdown selection
    drivers_list = [row['driver_name'] for row in conn.execute("SELECT driver_name FROM fleet_drivers").fetchall()]
    conn.close()
    
    col1, col2 = st.columns(2)
    with col1:
        ot_date = st.date_input("Select Target Date", value=date.today())
        date_str = ot_date.strftime("%Y-%m-%d")
        is_sunday = ot_date.weekday() == 6
        is_holiday = date_str in holiday_list
        
        # Rule Engine: Dynamic defaults for BOTH Start and End Times
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
        needs_transport = st.checkbox("Require Transportation Logistics?")
        
    with col2:
        if needs_transport:
            if drivers_list:
                selected_driver = st.selectbox("Select Driver Name", ["-- Select Driver --"] + drivers_list)
            else:
                selected_driver = "-- Select Driver --"
                st.warning("No drivers registered in fleet list.")

            # Auto-populate based on driver selection
            driver_name = ""
            driver_mobile = ""
            plate_number = ""
            
            if selected_driver != "-- Select Driver --":
                conn = get_db_connection()
                driver_info = conn.execute("SELECT * FROM fleet_drivers WHERE driver_name=?", (selected_driver,)).fetchone()
                conn.close()
                if driver_info:
                    driver_name = driver_info['driver_name']
                    driver_mobile = driver_info['driver_mobile']
                    plate_number = driver_info['plate_number']

            st.text_input("Driver Mobile Phone Number (Auto)", value=driver_mobile, disabled=True)
            st.text_input("Car Plate Registration Number (Auto)", value=plate_number, disabled=True)
            
            route_type = st.selectbox("Route Assignment Context", ["Weekday work", "Sunday work", "Sunday shopping", "Holiday Duty"])
            origin = st.text_input("Origin Address", value="Main Corporate Office")
            destination = st.text_input("Target Destination")
            dep_time = st.time_input("Departure Timeline Estimate", value=end_time)
            ret_time = st.time_input("Return Timeline Estimate", value=time(23, 0))
            
            # Pack times as string for form persistence
            dep_time_str = str(dep_time)
            ret_time_str = str(ret_time)
        else:
            driver_name, driver_mobile, plate_number, route_type, origin, destination, dep_time_str, ret_time_str = ["", "", "", "", "", "", "", ""]

    if st.button("Submit New Overtime Request"):
        conn = get_db_connection()
        conn.execute('''
            INSERT INTO overtime_requests (username, ot_date, start_time, end_time, needs_transport, 
            driver_name, driver_mobile, plate_number, route_type, origin, destination, departure_time, return_time)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (st.session_state.username, date_str, start_time.strftime("%H:%M"), end_time.strftime("%H:%M"),
              1 if needs_transport else 0, driver_name, driver_mobile, plate_number, route_type, origin, destination,
              dep_time_str, ret_time_str))
        
        user_info = conn.execute("SELECT email_recipients FROM users WHERE username=?", (st.session_state.username,)).fetchone()
        conn.commit()
        conn.close()
        
        st.success("🎉 Overtime and logistical logs successfully submitted!")
        if user_info and user_info['email_recipients']:
            st.info(f"📧 Notification pushed to predetermined dispatch recipients: **{user_info['email_recipients']}**")

    # Display Logs and Native Excel Export Interface
    st.subheader("📋 Overtime Submission History Log")
    
    conn = get_db_connection()
    if st.session_state.role == "Admin":
        ot_df = pd.read_sql_query("SELECT * FROM overtime_requests", conn)
    else:
        ot_df = pd.read_sql_query("SELECT * FROM overtime_requests WHERE username = ?", conn, params=[st.session_state.username])
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
        
        # ==============================================================================
        # 🚗 ADMIN ADVANCED OVERTIME EDITOR & FLEET ASSIGNMENT PANEL
        # ==============================================================================
        if st.session_state.role == "Admin":
            st.markdown("---")
            st.subheader("✏️ Edit Overtime Request & Fleet Assignment (Admin Only)")
            
            available_ids = ot_df['id'].tolist()
            selected_id = st.selectbox("Select Overtime Record ID to modify or assign transport:", available_ids)
            
            if selected_id:
                conn = get_db_connection()
                current_row = conn.execute("SELECT * FROM overtime_requests WHERE id = ?", (selected_id,)).fetchone()
                all_drivers = [row['driver_name'] for row in conn.execute("SELECT driver_name FROM fleet_drivers").fetchall()]
                conn.close()
                
                if current_row:
                    st.info(f"Modifying request submitted by personnel: **{current_row['username']}** for date **{current_row['ot_date']}**")
                    
                    with st.form(f"admin_edit_form_{selected_id}"):
                        edit_col1, edit_col2 = st.columns(2)
                        
                        with edit_col1:
                            try:
                                parsed_date = datetime.strptime(current_row['ot_date'], "%Y-%m-%d").date()
                            except:
                                parsed_date = date.today()
                                
                            edit_date = st.date_input("Overtime Date", value=parsed_date)
                            edit_start = st.text_input("OT Start Time (HH:MM)", value=current_row['start_time'])
                            edit_end = st.text_input("OT End Time (HH:MM)", value=current_row['end_time'])
                            edit_needs_trans = st.checkbox("Require Transportation Assignment?", value=bool(current_row['needs_transport']))
                            edit_origin = st.text_input("Origin Address", value=current_row['origin'] or "Main Corporate Office")
                            edit_destination = st.text_input("Target Destination", value=current_row['destination'] or "")

                        with edit_col2:
                            st.markdown("**Fleet & Vehicle Driver Assignment Desk:**")
                            if all_drivers:
                                try:
                                    d_idx = all_drivers.index(current_row['driver_name']) + 1
                                except:
                                    d_idx = 0
                                edit_driver_sel = st.selectbox("Assigned Driver Name", ["-- Select Driver --"] + all_drivers, index=d_idx)
                            else:
                                edit_driver_sel = "-- Select Driver --"
                            
                            edit_driver = edit_driver_sel if edit_driver_sel != "-- Select Driver --" else ""
                            edit_mobile = st.text_input("Driver Mobile Phone Number Override/Manual", value=current_row['driver_mobile'] or "")
                            edit_plate = st.text_input("Car Plate Registration Number Override/Manual", value=current_row['plate_number'] or "")
                            
                            route_options = ["Weekday work", "Sunday work", "Sunday shopping", "Holiday Duty"]
                            try:
                                default_idx = route_options.index(current_row['route_type'])
                            except:
                                default_idx = 0
                            edit_route = st.selectbox("Route Assignment Context", route_options, index=default_idx)
                            
                            edit_dep = st.text_input("Departure Time (HH:MM)", value=current_row['departure_time'] or "")
                            edit_ret = st.text_input("Return Time (HH:MM)", value=current_row['return_time'] or "")
                        
                        submit_changes = st.form_submit_button("Save Changes & Assign Fleet")
                        
                        if submit_changes:
                            # Auto-fill profile mapping if selected driver changed and overrides are not manual
                            if edit_driver_sel != "-- Select Driver --" and edit_mobile == "" and edit_plate == "":
                                conn = get_db_connection()
                                d_info = conn.execute("SELECT * FROM fleet_drivers WHERE driver_name=?", (edit_driver,)).fetchone()
                                conn.close()
                                if d_info:
                                    edit_mobile = d_info['driver_mobile']
                                    edit_plate = d_info['plate_number']

                            conn = get_db_connection()
                            conn.execute('''
                                UPDATE overtime_requests 
                                SET ot_date = ?, start_time = ?, end_time = ?, needs_transport = ?,
                                    driver_name = ?, driver_mobile = ?, plate_number = ?, route_type = ?,
                                    origin = ?, destination = ?, departure_time = ?, return_time = ?
                                WHERE id = ?
                            ''', (edit_date.strftime("%Y-%m-%d"), edit_start, edit_end, 1 if edit_needs_trans else 0,
                                  edit_driver, edit_mobile, edit_plate, edit_route,
                                  edit_origin, edit_destination, edit_dep, edit_ret, selected_id))
                            conn.commit()
                            conn.close()
                            
                            st.success(f"🎉 Overtime Request ID {selected_id} updated successfully!")
                            st.rerun()
    else:
        st.info("No recorded overtime history logs found.")


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
    else:
        st.info("No active meeting room allocations scheduled.")


# --- TAB 3: SYSTEM MASTER ADMINISTRATION CONTROL BOARDS ---
with tab3:
    if st.session_state.role != "Admin":
        st.error("🛡️ Restricted Access Control: You lack administrative clearing profiles to view these configuration matrices.")
    else:
        st.header("Admin Control Dashboard Engine")
        
        # ==============================================================================
        # 🗃️ MASTER DATABASE EXPLORER & CRUD SYSTEM (Admin Request)
        # ==============================================================================
        st.markdown("---")
        st.subheader("🗃️ Interactive Master Database Explorer (All Tables)")
        
        tables_list = ["users", "holidays", "overtime_requests", "meeting_rooms", "room_bookings", "fleet_drivers"]
        selected_table = st.selectbox("Select Table to View / Manage Data Records", tables_list)
        
        if selected_table:
            conn = get_db_connection()
            table_df = pd.read_sql_query(f"SELECT * FROM {selected_table}", conn)
            
            # Fetch primary keys / schema column lists
            cursor = conn.execute(f"PRAGMA table_info({selected_table})")
            columns_info = cursor.fetchall()
            columns_names = [col['name'] for col in columns_info]
            pk_col = [col['name'] for col in columns_info if col['pk'] == 1]
            pk_name = pk_col[0] if pk_col else columns_names[0]
            conn.close()
            
            st.markdown(f"**Live Records Grid view for table:** `{selected_table}`")
            st.dataframe(table_df, use_container_width=True)
            
            crud_action = st.radio("Choose Record Modification Task", ["Append New Row", "Update Existing Row", "Delete Row"])
            
            # 1. APPEND ROW FORM
            if crud_action == "Append New Row":
                st.markdown(f"**Append Data Entry into `{selected_table}`**")
                with st.form(f"append_form_{selected_table}"):
                    input_values = {}
                    for col in columns_names:
                        # Skip autoincrement ID row for additions
                        if col == 'id' and selected_table in ['overtime_requests', 'room_bookings']:
                            continue
                        input_values[col] = st.text_input(f"Enter `{col}` value")
                    
                    submit_append = st.form_submit_button("Append Data Record")
                    if submit_append:
                        conn = get_db_connection()
                        placeholders = ", ".join([f":{k}" for k in input_values.keys()])
                        fields = ", ".join(input_values.keys())
                        try:
                            conn.execute(f"INSERT INTO {selected_table} ({fields}) VALUES ({placeholders})", input_values)
                            conn.commit()
                            st.success(f"🎉 New record successfully appended to `{selected_table}`!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Execution Error: {str(e)}")
                        finally:
                            conn.close()
            
            # 2. UPDATE ROW FORM
            elif crud_action == "Update Existing Row":
                if table_df.empty:
                    st.info("Table is empty. Nothing to update.")
                else:
                    st.markdown(f"**Modify/Update Row Entry in `{selected_table}`**")
                    row_keys = table_df[pk_name].tolist()
                    row_to_update = st.selectbox(f"Select Target Identifier Row to Update ({pk_name})", row_keys)
                    
                    if row_to_update:
                        conn = get_db_connection()
                        current_entry = conn.execute(f"SELECT * FROM {selected_table} WHERE {pk_name} = ?", (row_to_update,)).fetchone()
                        conn.close()
                        
                        if current_entry:
                            with st.form(f"update_form_{selected_table}_{row_to_update}"):
                                updated_values = {}
                                for col in columns_names:
                                    if col == pk_name:
                                        st.text_input(f"`{col}` (Primary Key - Fixed)", value=str(current_entry[col]), disabled=True)
                                        continue
                                    updated_values[col] = st.text_input(f"Edit `{col}` value", value=str(current_entry[col] if current_entry[col] is not None else ""))
                                
                                submit_update = st.form_submit_button("Commit Updates")
                                if submit_update:
                                    conn = get_db_connection()
                                    update_set = ", ".join([f"{k} = :{k}" for k in updated_values.keys()])
                                    updated_values['pk_val'] = row_to_update
                                    try:
                                        conn.execute(f"UPDATE {selected_table} SET {update_set} WHERE {pk_name} = :pk_val", updated_values)
                                        conn.commit()
                                        st.success(f"🎉 Row `{row_to_update}` successfully updated in `{selected_table}`!")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Execution Error: {str(e)}")
                                    finally:
                                        conn.close()

            # 3. DELETE ROW INTERFACE
            elif crud_action == "Delete Row":
                if table_df.empty:
                    st.info("Table is empty. Nothing to remove.")
                else:
                    st.markdown(f"**Purge Data Record Row from `{selected_table}`**")
                    row_keys = table_df[pk_name].tolist()
                    row_to_delete = st.selectbox(f"Select Row Target Key to Delete ({pk_name})", row_keys)
                    
                    if st.button(f"Permanently Delete Record {row_to_delete}"):
                        conn = get_db_connection()
                        try:
                            conn.execute(f"DELETE FROM {selected_table} WHERE {pk_name} = ?", (row_to_delete,))
                            conn.commit()
                            st.success(f"💥 Record Row `{row_to_delete}` completely purged from `{selected_table}` database indices!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Execution Error: {str(e)}")
                        finally:
                            conn.close()
                            
        st.markdown("---")
