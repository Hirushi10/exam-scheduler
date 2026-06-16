import streamlit as st
import pandas as pd
from datetime import datetime, time
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import pymysql

st.set_page_config(page_title="Exams Scheduler", layout="wide")

# Custom CSS for the desktop app design
st.markdown("""
    <style>
    .main-title { font-size:26px; font-weight:bold; color: #1E3A8A; text-align: left; margin-bottom: 15px; border-bottom: 2px solid #1E3A8A; padding-bottom: 5px; }
    .hall-box { border: 1px solid #CBD5E1; padding: 15px; background-color: #FFFFFF; border-radius: 5px; margin-bottom: 15px; min-height: 220px; box-shadow: 1px 1px 5px rgba(0,0,0,0.05); }
    .hall-title { font-size: 16px; font-weight: bold; color: #1E40AF; border-bottom: 1px solid #E2E8F0; padding-bottom: 3px; margin-bottom: 8px; }
    .hall-details { font-size: 13px; color: #334155; line-height: 1.6; }
    .sidebar-section { background-color: #F8FAFC; padding: 15px; border-radius: 5px; border: 1px solid #E2E8F0; margin-bottom: 15px; }
    .summary-title { font-size: 15px; font-weight: bold; color: #0F172A; margin-bottom: 10px; border-bottom: 1px solid #CBD5E1; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<div class='main-title'>💻 Exams Scheduler - Faculty Of Management & Finance</div>", unsafe_allow_html=True)

# ----------------- DATABASE CONNECTION -----------------
def get_db_connection():
    return pymysql.connect(
        host=st.secrets["mysql"]["host"],
        port=int(st.secrets["mysql"]["port"]),
        database=st.secrets["mysql"]["database"],
        user=st.secrets["mysql"]["user"],
        password=st.secrets["mysql"]["password"],
        autocommit=True
    )

# Create table if it doesn't exist
def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS exam_assignments (
            id INT AUTO_INCREMENT PRIMARY KEY,
            semester_key VARCHAR(50),
            date VARCHAR(20),
            time VARCHAR(50),
            subject VARCHAR(255),
            hall VARCHAR(255),
            supervisor VARCHAR(255),
            invigilators TEXT
        )
    """)
    cursor.close()
    conn.close()

try:
    init_db()
except Exception as e:
    st.error(f"Database Connection/Initialization Error: {e}")
    st.stop()

# Helper function to load all active data from DB
def load_all_schedules():
    conn = get_db_connection()
    df = pd.read_sql("SELECT date AS Date, time AS Time, subject AS Subject, hall AS Hall, supervisor AS Supervisor, invigilators AS Invigilators, semester_key FROM exam_assignments", conn)
    conn.close()
    return df

schedule_df = load_all_schedules()

# Word Document Generation Function
def generate_word_report(df, title_text):
    doc = docx.Document()
    title = doc.add_paragraph()
    title_run = title.add_run(f"UNIVERSITY OF RUHUNA\nFaculty of Management and Finance\n{title_text}")
    title_run.bold = True
    title_run.font.size = Pt(14)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("\n")
    
    # Drop semester_key column for clean report if it exists
    report_df = df.drop(columns=['semester_key']) if 'semester_key' in df.columns else df
    
    table = doc.add_table(rows=1, cols=len(report_df.columns))
    table.style = 'Light Shading Accent 1'
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(report_df.columns):
        hdr_cells[i].text = str(col_name)
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        
    for index, row in report_df.iterrows():
        row_cells = table.add_row().cells
        for i, column in enumerate(report_df.columns):
            row_cells[i].text = str(row[column])
            
    filename = "Exam_Duty_Roster_Report.docx"
    doc.save(filename)
    return filename

# ----------------- LOAD EXCEL DATA (STAFF, HALLS, SUBJECTS) -----------------
try:
    staff_df = pd.read_excel("staff_list.xlsx")
    halls_df = pd.read_excel("halls_list.xlsx")
    subjects_df = pd.read_excel("subjects_list.xlsx")
except Exception as e:
    st.error(f"Excel Static Data Load Error: {e}")
    st.stop()

# ----------------- LIVE CALCULATIONS (LOAD SUMMARY) -----------------
duty_counts = {name: 0 for name in staff_df['Name'].tolist()}
if not schedule_df.empty:
    for _, row in schedule_df.iterrows():
        sup = row['Supervisor']
        if sup in duty_counts:
            duty_counts[sup] += 1
        invs_str = str(row['Invigilators'])
        if invs_str != "nan" and invs_str != "":
            for inv in [n.strip() for n in invs_str.split(",")]:
                if inv in duty_counts:
                    duty_counts[inv] += 1

summary_data = pd.DataFrame(list(duty_counts.items()), columns=['Staff Member', 'Duties']).sort_values(by='Duties', ascending=False)

# ----------------- LAYOUT SPLIT (SIDEBAR & MAIN) -----------------
col_control, col_display, col_summary = st.columns([1, 2.2, 0.8])

# COLUMN 1: CONTROLS & ASSIGNMENT (Left Side)
with col_control:
    st.markdown("### 🛠️ Control Panel")
    
    exam_date = st.date_input("Exam Date", datetime.now())
    date_str = str(exam_date)
    
    # Extract Year and Month Name for unique key configuration
    year_str = str(exam_date.year)
    month_str = exam_date.strftime("%B") # e.g., "June"
    
    # Semester Selection Dropdown
    semester_opt = st.selectbox("Select Semester", ["Sem 1", "Sem 2", "Sem 3"])
    
    # Automatically generate your unique partition key string
    generated_semester_key = f"{year_str}-{month_str}-{semester_opt.replace(' ', '')}"
    st.caption(f"Database Partition Key: `{generated_semester_key}`")
    
    start_time = st.time_input("Start Time", time(9, 0))
    end_time = st.time_input("End Time", time(12, 0))
    time_slot = f"{start_time.strftime('%I:%M %p')} - {end_time.strftime('%I:%M %p')}"
    st.info(f"Slot: {time_slot}")
    
    selected_sub = st.selectbox("Select Subject", subjects_df['Subject'].tolist())
    
    st.write("---")
    st.markdown("#### 📍 Assign Duty")
    selected_hall = st.selectbox("Select Target Hall", halls_df['Hall_Name'].tolist())
    
    # Clash check logic via live database entries
    busy_staff = []
    if not schedule_df.empty:
        current_clashes = schedule_df[(schedule_df['Date'] == date_str) & (schedule_df['Time'] == time_slot)]
        busy_staff += current_clashes['Supervisor'].tolist()
        for inv_list_str in current_clashes['Invigilators'].tolist():
            if pd.notna(inv_list_str) and inv_list_str != "" and inv_list_str != "nan":
                busy_staff += [n.strip() for n in inv_list_str.split(",")]

    available_supervisors = staff_df[(staff_df['Role'] == 'Supervisor') & (~staff_df['Name'].isin(busy_staff))]['Name'].tolist()
    available_invigilators = staff_df[(staff_df['Role'] == 'Invigilator') & (~staff_df['Name'].isin(busy_staff))]['Name'].tolist()

    selected_sup = st.selectbox("Controller (Supervisor)", ["-- Select --"] + available_supervisors)
    selected_invs = st.multiselect("Invigilators", available_invigilators)
    
    if st.button("💾 Save Assignment", use_container_width=True, type="primary"):
        if selected_sup != "-- Select --" and len(selected_invs) > 0:
            if not schedule_df.empty and not schedule_df[(schedule_df['Date'] == date_str) & (schedule_df['Time'] == time_slot) & (schedule_df['Hall'] == selected_hall)].empty:
                st.error(f"{selected_hall} is already allocated for this slot!")
            else:
                invs_string = ", ".join(selected_invs)
                
                try:
                    conn = get_db_connection()
                    cursor = conn.cursor()
                    query = """
                        INSERT INTO exam_assignments (semester_key, date, time, subject, hall, supervisor, invigilators)
                        VALUES (%s, %s, %s, %s, %s, %s, %s)
                    """
                    cursor.execute(query, (generated_semester_key, date_str, time_slot, selected_sub, selected_hall, selected_sup, invs_string))
                    cursor.close()
                    conn.close()
                    
                    st.success("Saved securely to Cloud MySQL database!")
                    st.rerun()
                except Exception as ex:
                    st.error(f"DB Insert Error: {ex}")
        else:
            st.error("Select both Controller and Invigilators.")
            
    st.write("---")
    st.markdown("#### 📥 Current Roster Report")
    
    # Roster output directly mapped from the database snapshot
    if not schedule_df.empty:
        file_path = generate_word_report(schedule_df, "EXAMINATION DUTY ROSTER")
        with open(file_path, "rb") as file:
            st.download_button(
                label="📥 Download Word Report",
                data=file,
                file_name=f"Faculty_Exam_Roster_{generated_semester_key}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
    else:
        st.button("📥 Download Word Report", disabled=True, use_container_width=True)

# COLUMN 2: VISUAL HALL BOXES / CARDS (Middle)
with col_display:
    st.markdown(f"### 🏢 Available Exam Halls View ({date_str} | {time_slot})")
    
    active_slots = schedule_df[(schedule_df['Date'] == date_str) & (schedule_df['Time'] == time_slot)] if not schedule_df.empty else pd.DataFrame()
    all_halls = halls_df['Hall_Name'].tolist()
    grid_cols = st.columns(2)
    
    for idx, hall_name in enumerate(all_halls):
        target_col = grid_cols[idx % 2]
        with target_col:
            hall_record = active_slots[active_slots['Hall'] == hall_name] if not active_slots.empty else pd.DataFrame()
            if not hall_record.empty:
                sup_assigned = hall_record.iloc[0]['Supervisor']
                invs_assigned = hall_record.iloc[0]['Invigilators']
                sub_assigned = hall_record.iloc[0]['Subject']
                date_assigned = hall_record.iloc[0]['Date']
                time_assigned = hall_record.iloc[0]['Time']
                
                st.markdown(f"""
                    <div class="hall-box" style="border-left: 5px solid #10B981; background-color: #F0FDF4;">
                        <div class="hall-title">🟢 {hall_name}</div>
                        <div class="hall-details">
                            <b>Schedule:</b> {date_assigned} ({time_assigned})<br>
                            <b>Subject:</b> {sub_assigned}<br>
                            <b>Supervisor:</b> {sup_assigned}<br>
                            <b>Invigilators:</b> {invs_assigned}
                        </div>
                    </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                    <div class="hall-box" style="border-left: 5px solid #64748B; background-color: #F8FAFC;">
                        <div class="hall-title" style="color:#64748B;">⚪ {hall_name}</div>
                        <div class="hall-details" style="color:#94A3B8;">
                            <i>No exam scheduled or staff assigned for this time slot.</i>
                        </div>
                    </div>
                """, unsafe_allow_html=True)

# COLUMN 3: LOAD SUMMARY LIST (Right Side)
with col_summary:
    st.markdown("### 📊 Summary")
    st.markdown("<div class='summary-title'>Load Summary</div>", unsafe_allow_html=True)
    st.dataframe(summary_data, use_container_width=True, hide_index=True, height=520)

st.write("---")
# ----------------- HISTORICAL DATABASE SEARCH VIEW PANEL -----------------
st.markdown("### 🔍 Historical Semester Archives Look-up Panel")
if not schedule_df.empty:
    unique_keys = sorted(schedule_df['semester_key'].dropna().unique())
    selected_archive_key = st.selectbox("Select Historical Key Block to Fetch", unique_keys)
    
    filtered_history = schedule_df[schedule_df['semester_key'] == selected_archive_key]
    
    st.dataframe(filtered_history.drop(columns=['semester_key']), use_container_width=True, hide_index=True)
    
    history_file_path = generate_word_report(filtered_history, f"EXAMINATION DUTY ROSTER - ARCHIVE [{selected_archive_key}]")
    with open(history_file_path, "rb") as file_hist:
        st.download_button(
            label=f"📥 Download Word Report for {selected_archive_key}",
            data=file_hist,
            file_name=f"Archive_Exam_Roster_{selected_archive_key}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.info("No saved records inside the database yet to display logs.")
