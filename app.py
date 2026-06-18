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
    .hall-box { border: 1px solid #CBD5E1; padding: 15px; background-color: #FFFFFF; border-radius: 5px; margin-bottom: 15px; min-height: 240px; box-shadow: 1px 1px 5px rgba(0,0,0,0.05); }
    .hall-title { font-size: 16px; font-weight: bold; color: #1E40AF; border-bottom: 1px solid #E2E8F0; padding-bottom: 3px; margin-bottom: 8px; }
    .hall-details { font-size: 13px; color: #334155; line-height: 1.6; }
    .sidebar-section { background-color: #F8FAFC; padding: 15px; border-radius: 5px; border: 1px solid #E2E8F0; margin-bottom: 15px; }
    .summary-title { font-size: 15px; font-weight: bold; color: #0F172A; margin-bottom: 10px; border-bottom: 1px solid #CBD5E1; }
    .badge-capacity { font-size: 11px; font-weight: bold; background-color: #E2E8F0; color: #475569; padding: 3px 6px; border-radius: 4px; float: right; }
    .badge-danger { font-size: 11px; font-weight: bold; background-color: #EF4444; color: #FFFFFF; padding: 3px 6px; border-radius: 4px; float: right; }
    .badge-success { font-size: 11px; font-weight: bold; background-color: #10B981; color: #FFFFFF; padding: 3px 6px; border-radius: 4px; float: right; }
    
    /* Force Hand Pointer Symbol on Dropdowns, Inputs, and Selections */
    div[data-baseweb="select"] { cursor: pointer !important; }
    div[data-baseweb="select"] * { cursor: pointer !important; }
    div[data-testid="stDateInput"] input { cursor: pointer !important; }
    div[data-testid="stTimeInput"] input { cursor: pointer !important; }
    button { cursor: pointer !important; }
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
        autocommit=True,
        connect_timeout=10
    )

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
    
    cursor.execute("SHOW COLUMNS FROM exam_assignments LIKE 'student_count'")
    if not cursor.fetchone():
        cursor.execute("ALTER TABLE exam_assignments ADD COLUMN student_count INT DEFAULT 0")
        
    cursor.close()
    conn.close()

try:
    init_db()
except Exception as e:
    st.error(f"Database Initialization Error: {e}")
    st.stop()

def load_all_schedules():
    conn = get_db_connection()
    query = "SELECT id, date AS Date, time AS Time, subject AS Subject, hall AS Hall, supervisor AS Supervisor, invigilators AS Invigilators, semester_key, student_count FROM exam_assignments"
    df = pd.read_sql(query, conn)
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
    
    cols_to_drop =
