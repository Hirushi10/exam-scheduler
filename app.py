import streamlit as st
import pandas as pd
from datetime import datetime
import sqlite3

# --- 1. DATABASE SETUP & INITIALIZATION ---
def init_db():
    conn = sqlite3.connect('semexam.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS exam_assignments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            semester_key TEXT,
            date TEXT,
            time TEXT,
            subject TEXT,
            hall TEXT,
            supervisor TEXT,
            invigilators TEXT,
            student_count INTEGER DEFAULT 0
        )
    ''')
    # Check if student_count column exists (migration safety)
    cursor.execute("PRAGMA table_info(exam_assignments)")
    columns = [col[1] for col in cursor.execute("PRAGMA table_info(exam_assignments)").fetchall()]
    if 'student_count' not in columns:
        cursor.execute("ALTER TABLE exam_assignments ADD COLUMN student_count INTEGER DEFAULT 0")
    conn.commit()
    conn.close()

init_db()

# --- 2. EXCEL DATASETS PARSING ---
@st.cache_data(ttl=10)  # Refreshes cache periodically if clerks modify Excel sheets
def load_excel_data():
    try:
        # Load Staff list
        staff_df = pd.read_excel('staff_list.xlsx')
        staff_list = staff_df.iloc[:, 0].dropna().str.strip().tolist()
        
        # Load Halls list with Capacity parsing (Column A: Name, Column B: Max Capacity)
        halls_df = pd.read_excel('halls_list.xlsx')
        halls_list = halls_df.iloc[:, 0].dropna().str.strip().tolist()
        
        hall_capacities = {}
        for _, row in halls_df.iterrows():
            if pd.notna(row.iloc[0]):
                h_name = str(row.iloc[0]).strip()
                # Default to N/A if column B doesn't have numerical integers
                h_cap = int(row.iloc[1]) if len(row) > 1 and pd.notna(row.iloc[1]) else "N/A"
                hall_capacities[h_name] = h_cap
                
        # Load Subjects list
        subjects_df = pd.read_excel('subjects_list.xlsx')
        subjects_list = subjects_df.iloc[:, 0].dropna().str.strip().tolist()
        
        return staff_list, halls_list, hall_capacities, subjects_list
    except Exception as e:
        st.error(f"Excel Resource Error: Make sure lists are placed inside the folder. Details: {e}")
        return [], [], {}, []

all_staff, all_halls, hall_capacities, all_subjects = load_excel_data()

# --- 3. HELPER DATABASE ENGINE ACTIONS ---
def fetch_all_assignments():
    conn = sqlite3.connect('semexam.db')
    df = pd.read_sql_query("SELECT * FROM exam_assignments ORDER BY id DESC", conn)
    conn.close()
    return df

def delete_assignment(assignment_id):
    conn = sqlite3.connect('semexam.db')
    cursor = conn.cursor()
    cursor.execute("DELETE FROM exam_assignments WHERE id = ?", (assignment_id,))
    conn.commit()
    conn.close()

# --- 4. STREAMLIT UI VIEW PORT & WORKFLOW ---
st.set_page_config(page_title="Exams Scheduler", layout="wide")
st.title("💻 Exams Scheduler - Faculty Of Management & Finance")

# Session state trackers for state management and handling edit rollbacks smoothly
if 'edit_id' not in st.session_state:
    st.session_state.edit_id = None
if 'form_data' not in st.session_state:
    st.session_state.form_data = {}

# Load current raw assignments
assignments_df = fetch_all_assignments()

# --- 5. WORKLOAD STATISTICS MATRICES ---
duty_counts = {name: 0 for name in all_staff}
for _, row in assignments_df.iterrows():
    sup = row['supervisor']
    if sup in duty_counts:
        duty_counts[sup] += 1
    if row['invigilators']:
        invs = [i.strip() for i in str(row['invigilators']).split(',') if i.strip()]
        for inv in invs:
            if inv in duty_counts:
                duty_counts[inv] += 1

# --- 6. CORE APP FIELDS & CONTROLS INTERLOCK ---
col1, col2, col3 = st.columns([1, 2, 1])

with col1:
    st.subheader("🛠️ Control Panel" if not st.session_state.edit_id else "📝 Edit Mode Active")
    
    # Defaults checking if edit mode context parameters are triggered
    default_date = datetime.now()
    default_sem = "Sem 1"
    default_sub = all_subjects[0] if all_subjects else ""
    default_hall = all_halls[0] if all_halls else ""
    default_sup = ""
    default_invs = []
    default_students = 0
    
    if st.session_state.edit_id and not assignments_df.empty:
        target = assignments_df[assignments_df['id'] == st.session_state.edit_id]
        if not target.empty:
            default_date = datetime.strptime(target.iloc[0]['date'], '%Y-%m-%d')
            # Extract time structures back to slots
            time_parts = target.iloc[0]['time'].split(' - ')
            default_sub = target.iloc[0]['subject']
            default_hall = target.iloc[0]['hall']
            default_sup = target.iloc[0]['supervisor']
            default_invs = [i.strip() for i in str(target.iloc[0]['invigilators']).split(',') if i.strip()]
            default_students = int(target.iloc[0]['student_count'])

    exam_date = st.date_input("Exam Date", value=default_date)
    semester = st.selectbox("Select Semester", ["Sem 1", "Sem 2", "Sem 3"])
    
    t_col1, t_col2 = st.columns(2)
    with t_col1:
        start_time = st.time_input("Start Time", value=datetime.strptime("09:00", "%H:%M").time())
    with t_col2:
        end_time = st.time_input("End Time", value=datetime.strptime("12:00", "%H:%M").time())
        
    time_slot_string = f"{start_time.strftime('%I:%M %p')} - {end_time.strftime('%I:%M %p')}"
    st.info(f"Slot: {time_slot_string}")
    
    subject = st.selectbox("Select Subject", all_subjects, index=all_subjects.index(default_sub) if default_sub in all_subjects else 0)
    
    # 🔢 STUDENT COUNT INPUT ELEMENT
    student_count = st.number_input("Expected Student Count", min_value=0, value=default_students, step=1)
    
    st.markdown("---")
    st.write("##### 📍 Assign Duty")
    target_hall = st.selectbox("Select Target Hall", all_halls, index=all_halls.index(default_hall) if default_hall in all_halls else 0)

    # Calculate conflicting/busy staff allocations live for this slot parameters
    busy_staff = []
    active_slot_mapping = {}
    
    for _, row in assignments_df.iterrows():
        if row['date'] == str(exam_date) and row['time'] == time_slot_string:
            active_slot_mapping[row['hall']] = row
            
            # Skip current record if in edit routing rules to prevent trapping itself
            if st.session_state.edit_id and row['id'] == st.session_state.edit_id:
                continue
                
            busy_staff.append(row['supervisor'])
            if row['invigilators']:
                busy_staff.extend([i.strip() for i in str(row['invigilators']).split(',') if i.strip()])
                
    busy_staff = list(set(busy_staff))
    available_staff_pool = [name for name in all_staff if name not in busy_staff]

    # Interlock logic safety: Force re-injection if default targets are assigned during editing
    if st.session_state.edit_id:
        if default_sup and default_sup not in available_staff_pool:
            available_staff_pool.append(default_sup)
        for inv in default_invs:
            if inv not in available_staff_pool:
                available_staff_pool.append(inv)

    # Supervisor Single Select drop-down filter
    sup_options = ["-- Select --"] + available_staff_pool
    sup_index = sup_options.index(default_sup) if default_sup in sup_options else 0
    supervisor = st.selectbox("Controller (Supervisor)", sup_options, index=sup_index)

    # Invigilators Multi-Select filters (Excludes selected supervisor dynamically)
    inv_pool = [name for name in available_staff_pool if name != supervisor]
    invigilators = st.multiselect("Invigilators (Select Multiple)", inv_pool, default=[i for i in default_invs if i in inv_pool])

    # Save button operations routing rules
    if st.button("💾 Update Assignment" if st.session_state.edit_id else "💾 Save Assignment", type="primary", use_container_width=True):
        if supervisor == "-- Select --" or not invigilators:
            st.error("Validation Clash: You must specify a Supervisor and at least one Invigilator.")
        elif target_hall in active_slot_mapping and (not st.session_state.edit_id or active_slot_mapping[target_hall]['id'] != st.session_state.edit_id):
            st.error(f"Conflict Crash: {target_hall} is already assigned a dynamic roster block for this specific slot!")
        else:
            # Generate partition keys
            dt = datetime.combine(exam_date, datetime.min.time())
            semester_key = f"{dt.strftime('%Y-%B')}-{semester.replace(' ', '')}"
            invs_string = ", ".join(invigilators)
            
            conn = sqlite3.connect('semexam.db')
            cursor = conn.cursor()
            if st.session_state.edit_id:
                cursor.execute('''
                    UPDATE exam_assignments 
                    SET semester_key=?, date=?, time=?, subject=?, hall=?, supervisor=?, invigilators=?, student_count=?
                    WHERE id=?
                ''', (semester_key, str(exam_date), time_slot_string, subject, target_hall, supervisor, invs_string, student_count, st.session_state.edit_id))
                st.session_state.edit_id = None # Break edit lock state registers
            else:
                cursor.execute('''
                    INSERT INTO exam_assignments (semester_key, date, time, subject, hall, supervisor, invigilators, student_count)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (semester_key, str(exam_date), time_slot_string, subject, target_hall, supervisor, invs_string, student_count))
            conn.commit()
            conn.close()
            st.rerun()
            
    if st.session_state.edit_id:
        if st.button("Cancel Edit", use_container_width=True):
            st.session_state.edit_id = None
            st.rerun()

# --- 7. COLUMN 2: AVAILABLE EXAM HALLS STATUS VIEW CARDS GRID ---
with col2:
    st.subheader(f"🏢 Available Exam Halls View Grid ({exam_date} | {time_slot_string})")
    
    h_col1, h_col2 = st.columns(2)
    for index, hall_name in enumerate(all_halls):
        target_column = h_col1 if index % 2 == 0 else h_col2
        max_capacity = hall_capacities.get(hall_name, "N/A")
        
        with target_column:
            if hall_name in active_slot_mapping:
                record = active_slot_mapping[hall_name]
                current_students = record['student_count']
                
                # Check for capacity breach over-allocation warnings
                is_overcrowded = max_capacity != "N/A" and current_students > max_capacity
                
                if is_overcrowded:
                    st.error(f"⚠️ **{hall_name}** [Students: {current_students} / {max_capacity} Max]\n\n"
                             f"**Subject:** {record['subject']}\n\n"
                             f"**Supervisor:** {record['supervisor']}\n\n"
                             f"**Invigilators:** {record['invigilators']}")
                else:
                    st.success(f"🟢 **{hall_name}** [Students: {current_students} / {max_capacity} Max]\n\n"
                               f"**Subject:** {record['subject']}\n\n"
                               f"**Supervisor:** {record['supervisor']}\n\n"
                               f"**Invigilators:** {record['invigilators']}")
            else:
                st.info(f"⚪ **{hall_name}** [Max Capacity: {max_capacity}]\n\n*No exam scheduled or staff assigned for this time slot.*")

# --- 8. COLUMN 3: REAL-TIME TOTAL WORKLOADCOUNTER MATRIX ---
with col3:
    st.subheader("📊 Workload Matrix")
    summary_df = pd.DataFrame(list(duty_counts.items()), columns=['Staff Member', 'Duties']).sort_values(by='Duties', ascending=False)
    st.dataframe(summary_df, use_container_width=True, hide_index=True, height=400)

st.markdown("---")

# --- 9. ACTIVE ASSIGNMENTS MANAGEMENT PANEL (LIVE EDITOR) ---
st.subheader("📝 Active Assignments Management Panel (Live Editor)")
if not assignments_df.empty:
    for _, row in assignments_df.iterrows():
        with st.container():
            m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns([1.5, 1.5, 2, 3, 1.5])
            with m_col1:
                st.write(f"📅 **{row['date']}**")
                st.caption(f"ID: {row['id']} | {row['semester_key']}")
            with m_col2:
                st.write(f"⏰ {row['time']}")
            with m_col3:
                st.write(f"🏢 **{row['hall']}** (Allocated: {row['student_count']})")
                st.caption(f"📚 {row['subject']}")
            with m_col4:
                st.write(f"👤 **Sup:** {row['supervisor']}")
                st.write(f"👥 **Invs:** {row['invigilators']}")
            with m_col5:
                e_btn, d_btn = st.columns(2)
                with e_btn:
                    if st.button("✏️", key=f"edit_{row['id']}", use_container_width=True):
                        st.session_state.edit_id = row['id']
                        st.rerun()
                with d_btn:
                    if st.button("🗑️", key=f"del_{row['id']}", use_container_width=True):
                        delete_assignment(row['id'])
                        st.success(f"Log row identity record cleared!")
                        st.rerun()
            st.markdown("<hr style='margin:4px 0px; border-color:#eee;'>", unsafe_allow_width=True)
else:
    st.info("No active logs discovered inside the system roster files.")
