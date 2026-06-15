import streamlit as st
import pandas as pd
from datetime import datetime, time
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

st.set_page_config(page_title="M Exams Scheduler", layout="wide")

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

# Word Document Generation Function
def generate_word_report(df, title_text):
    doc = docx.Document()
    title = doc.add_paragraph()
    title_run = title.add_run(f"UNIVERSITY OF RUHUNA\nFaculty of Management and Finance\n{title_text}")
    title_run.bold = True
    title_run.font.size = Pt(14)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("\n")
    
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = 'Light Shading Accent 1'
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df.columns):
        hdr_cells[i].text = str(col_name)
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        
    for index, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, column in enumerate(df.columns):
            row_cells[i].text = str(row[column])
            
    filename = "Exam_Duty_Roster_Report.docx"
    doc.save(filename)
    return filename

# ----------------- LOAD EXCEL DATA -----------------
try:
    staff_df = pd.read_excel("staff_list.xlsx")
    halls_df = pd.read_excel("halls_list.xlsx")
    subjects_df = pd.read_excel("subjects_list.xlsx")
    if os.path.exists("current_schedule.xlsx") and os.path.getsize("current_schedule.xlsx") > 0:
        schedule_df = pd.read_excel("current_schedule.xlsx")
        schedule_df['Date'] = schedule_df['Date'].astype(str)
        schedule_df['Invigilators'] = schedule_df['Invigilators'].astype(str)
    else:
        schedule_df = pd.DataFrame(columns=['Date', 'Time', 'Subject', 'Hall', 'Supervisor', 'Invigilators'])
except Exception as e:
    st.error(f"Excel Data Load Error: {e}")
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
    
    start_time = st.time_input("Start Time", time(9, 0))
    end_time = st.time_input("End Time", time(12, 0))
    time_slot = f"{start_time.strftime('%I:%M %p')} - {end_time.strftime('%I:%M %p')}"
    st.info(f"Slot: {time_slot}")
    
    selected_sub = st.selectbox("Select Subject", subjects_df['Subject'].tolist())
    
    st.write("---")
    st.markdown("#### 📍 Assign Duty")
    selected_hall = st.selectbox("Select Target Hall", halls_df['Hall_Name'].tolist())
    
    # Clash check logic for the selected slot
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
                new_row = {
                    'Date': date_str, 'Time': time_slot, 'Subject': selected_sub, 
                    'Hall': selected_hall, 'Supervisor': selected_sup, 'Invigilators': invs_string
                }
                try:
                    schedule_df = pd.concat([schedule_df, pd.DataFrame([new_row])], ignore_index=True)
                    schedule_df.to_excel("current_schedule.xlsx", index=False)
                    st.success("Saved successfully!")
                    st.rerun()
                except PermissionError:
                    st.error("Close current_schedule.xlsx first!")
        else:
            st.error("Select both Controller and Invigilators.")
            
    st.write("---")
    st.markdown("#### 📥 Reports & Reset")
    
    # Word Report Button added to Control Panel
    if not schedule_df.empty:
        file_path = generate_word_report(schedule_df, "EXAMINATION DUTY ROSTER")
        with open(file_path, "rb") as file:
            st.download_button(
                label="📥 Download Word Report",
                data=file,
                file_name="Faculty_Exam_Roster.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
    else:
        st.button("📥 Download Word Report", disabled=True, use_container_width=True)
            
    if st.button("🗑️ Clean All Schedule", type="secondary", use_container_width=True):
        schedule_df = pd.DataFrame(columns=['Date', 'Time', 'Subject', 'Hall', 'Supervisor', 'Invigilators'])
        schedule_df.to_excel("current_schedule.xlsx", index=False)
        st.rerun()

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
    st.dataframe(summary_data, use_container_width=True, hide_index=True, height=550)
