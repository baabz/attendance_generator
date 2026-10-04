"""
Streamlit Web Application: Exam Hall Allocation & Attendance Sheet Generator
Enterprise-grade, polished UI tailored for Kaduna Polytechnic and adaptable across all departments.
Features professional vector iconography (Material Symbols) and clean corporate typography.
"""
import datetime
import importlib
import io
import json
import os
import random
import pandas as pd
import streamlit as st

def _ensure_static_assets():
    try:
        import streamlit as _st
        js_dir = os.path.join(os.path.dirname(_st.__file__), "static", "static", "js")
        tc_path = os.path.join(js_dir, "TableChart.esm.o-QINulw.js")
        if os.path.exists(tc_path):
            with open(tc_path, "rb") as f:
                data = f.read()
            if b"\x00" in data or len(data) == 0:
                with open(tc_path, "w", encoding="utf-8") as f:
                    f.write('import{o as e}from"./rolldown-runtime.C0FnF6B9.js";import{r as t}from"./emotion-is-prop-valid.esm.CuYlOKOk.js";import{v as n}from"./emotion-styled.browser.esm.Dj5rZcu_.js";import{t as r}from"./index.esm.M9ThdiDn.js";var i=e(t()),a=i.forwardRef(function(e,t){return i.createElement(r,n({iconAttrs:{fill:`currentColor`,xmlns:`http://www.w3.org/2000/svg`},iconVerticalAlign:`middle`,iconViewBox:`0 0 24 24`},e,{ref:t}),i.createElement(`path`,{fill:`none`,d:`M0 0h24v24H0V0z`}),i.createElement(`path`,{d:`M10 10.02h5V21h-5zM17 21h3c1.1 0 2-.9 2-2v-9h-5v11zm3-18H5c-1.1 0-2 .9-2 2v3h19V5c0-1.1-.9-2-2-2zM3 19c0 1.1.9 2 2 2h3V10H3v9z`}))});a.displayName=`TableChart`;export{a as t};')
    except Exception:
        pass

_ensure_static_assets()

import core.models
import core.parser
import core.shuffler
import core.docx_generator
import core.pdf_generator
import core.door_notices
import core.envelope_tags

importlib.reload(core.models)
importlib.reload(core.parser)
importlib.reload(core.shuffler)
importlib.reload(core.docx_generator)
importlib.reload(core.pdf_generator)
importlib.reload(core.door_notices)
importlib.reload(core.envelope_tags)

from core.models import Student, ExamConfig, VenueAllocation
from core.parser import (
    parse_excel,
    parse_csv,
    parse_docx,
    parse_raw_text,
    extract_file_to_dataframe,
    parse_dataframe,
    generate_blank_excel_template,
    detect_columns
)
from core.shuffler import allocate_to_venues
from core.docx_generator import build_attendance_docx
from core.pdf_generator import build_attendance_pdf
from core.door_notices import build_door_notices_docx, build_door_notices_pdf
from core.envelope_tags import build_envelope_tags_docx, build_envelope_tags_pdf


# Page setup
st.set_page_config(
    page_title="Kaduna Polytechnic — Examination Seating & Attendance Register",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Professional Enterprise Theme CSS with Material Symbols
st.markdown("""
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200" />
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Material Symbols Utility */
    .material-symbols-outlined {
        font-size: 19px;
        vertical-align: -3px;
        line-height: 1;
        display: inline-block;
    }
    .icon-brand {
        color: #1E4E79;
    }
    .icon-white {
        color: #FFFFFF;
    }
    .icon-success {
        color: #15803D;
    }
    .icon-warn {
        color: #A16207;
    }
    .icon-error {
        color: #B91C1C;
    }
    
    /* Hero Top Banner */
    .hero-banner {
        background: linear-gradient(135deg, #0A2239 0%, #153E65 50%, #1E4E79 100%);
        color: white;
        padding: 22px 28px;
        border-radius: 8px;
        margin-bottom: 22px;
        box-shadow: 0 2px 8px rgba(10, 34, 57, 0.1);
        border-left: 5px solid #2B6CB0;
    }
    .hero-badge {
        display: inline-block;
        background: rgba(255, 255, 255, 0.15);
        padding: 3px 10px;
        border-radius: 4px;
        font-size: 11px;
        font-weight: 600;
        letter-spacing: 0.6px;
        text-transform: uppercase;
        margin-bottom: 6px;
    }
    .hero-title {
        font-size: 22px;
        font-weight: 700;
        margin: 0;
        letter-spacing: -0.2px;
        color: #FFFFFF;
    }
    .hero-subtitle {
        font-size: 13px;
        opacity: 0.88;
        margin-top: 5px;
        margin-bottom: 0;
        font-weight: 400;
    }
    
    /* Section Headings */
    .section-title {
        font-size: 14.5px;
        font-weight: 700;
        color: #1E293B;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    /* Cards */
    [data-testid="stVerticalBlockBorderWrapper"] {
        border-color: #E2E8F0 !important;
        background-color: #FFFFFF;
        border-radius: 8px !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
    }
    
    /* Metrics Pill */
    .metric-pill {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        padding: 9px 14px;
        border-radius: 8px;
        text-align: center;
    }
    .metric-pill-label {
        font-size: 11px;
        color: #64748B;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.3px;
    }
    .metric-pill-value {
        font-size: 18px;
        font-weight: 700;
        color: #0F172A;
        margin-top: 2px;
    }
    
    /* Buttons */
    .stButton > button {
        border-radius: 6px;
        font-weight: 600;
        font-size: 13.5px;
        transition: all 0.2s ease;
    }
    
    /* Tab Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        border-bottom: 1px solid #E2E8F0;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 7px 16px;
        font-size: 13px;
        font-weight: 500;
    }
</style>
""", unsafe_allow_html=True)


def load_presets():
    preset_path = os.path.join(os.path.dirname(__file__), "config", "presets.json")
    if os.path.exists(preset_path):
        with open(preset_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


presets = load_presets()

# Institutional Header
st.markdown("""
<div class="hero-banner">
    <h1 class="hero-title">Examination Seating & Attendance Register Portal</h1>
    <p class="hero-subtitle">Official examination materials preparation, hall capacity allocation, and security register generator.</p>
</div>
""", unsafe_allow_html=True)


# Initialize Session State
if "venues_list" not in st.session_state:
    st.session_state["venues_list"] = []

if "venue_id_counter" not in st.session_state:
    st.session_state["venue_id_counter"] = 1

if "students_list" not in st.session_state:
    st.session_state["students_list"] = []

if "custom_seed" not in st.session_state:
    st.session_state["custom_seed"] = f"DAY1_{datetime.date.today().strftime('%b%d').upper()}"


# Primary Two-Column Layout
col_left, col_right = st.columns([1, 1], gap="medium")


# ==========================================
# LEFT COLUMN: Exam Details & Student List
# ==========================================
with col_left:
    # Card 1: Exam Details
    with st.container(border=True):
        st.markdown(
            '<div class="section-title"><span class="material-symbols-outlined icon-brand">description</span> Section 1: Course & Examination Specifications</div>',
            unsafe_allow_html=True
        )
        
        dept_options = presets.get("departments", [])
        selected_dept = st.selectbox(
            "Academic Department",
            options=["-- Select Department --"] + dept_options + ["Custom (Type Manually)"],
            index=0
        )
        if selected_dept == "Custom (Type Manually)":
            dept_name = st.text_input("Enter Department Name", value="", placeholder="e.g. DEPARTMENT OF COMPUTER ENGINEERING")
        elif selected_dept == "-- Select Department --":
            dept_name = ""
        else:
            dept_name = selected_dept

        c_col1, c_col2 = st.columns(2)
        with c_col1:
            college_name = st.text_input(
                "College / Institution",
                value=presets.get("college", "KADUNA POLYTECHNIC"),
                placeholder="e.g. KADUNA POLYTECHNIC"
            )
            course_title = st.text_input(
                "Course Title",
                value="",
                placeholder="e.g. ELECTRICAL ENGINEERING SCIENCE II"
            )
            level_options = ["-- Select Level --"] + presets.get("levels", ["ND I", "ND II", "HND I", "HND II"])
            selected_level = st.selectbox(
                "Academic Level",
                options=level_options,
                index=0
            )
            level = "" if selected_level == "-- Select Level --" else selected_level

        with c_col2:
            session_options = presets.get("sessions", ["2nd/4th SEMESTER, 2025/2026 SESSION (REGULAR)"])
            session_val = st.selectbox("Academic Session & Semester", options=session_options + ["Custom..."])
            if session_val == "Custom...":
                session_val = st.text_input("Custom Session/Semester", value="", placeholder="e.g. 2nd/4th SEMESTER, 2025/2026 SESSION (REGULAR)")
            course_code = st.text_input(
                "Course Code",
                value="",
                placeholder="e.g. EEC 239 or CTE 443"
            )
            exam_date = st.text_input(
                "Examination Date",
                value="",
                placeholder="e.g. 15th October, 2026"
            )

        with st.expander("Timing & Invigilation Endorsement Specifications", expanded=False):
            t_col1, t_col2 = st.columns(2)
            with t_col1:
                time_in = st.text_input("Examination Start Time", value="", placeholder="e.g. 09:00 AM")
                chief_invigilator = st.text_input("Chief Invigilator Designation Line", value="______________________________________")
            with t_col2:
                time_out = st.text_input("Examination End Time", value="", placeholder="e.g. 12:00 PM")
                invig_count = st.number_input("Invigilator Signature Rows", min_value=1, max_value=6, value=3)

            continuous_sn = st.checkbox(
                "Continuous Candidate Serial Numbering (Across All Venues)",
                value=False,
                help="By default, S/N restarts at 1 for each venue so invigilators can verify attendance count at a glance."
            )

    exam_config = ExamConfig(
        session_semester=session_val,
        department=dept_name,
        college=college_name,
        course_title=course_title,
        course_code=course_code,
        exam_date=exam_date,
        time_in=time_in,
        time_out=time_out,
        level=level,
        chief_invigilator=chief_invigilator,
        invigilator_count=invig_count,
        continuous_sn=continuous_sn
    )

    # Card 2: Student Master List
    with st.container(border=True):
        st.markdown(
            '<div class="section-title"><span class="material-symbols-outlined icon-brand">groups</span> Section 2: Candidate Master Register</div>',
            unsafe_allow_html=True
        )
        
        tab_upload, tab_paste, tab_sample = st.tabs([
            "Upload Roster File",
            "Direct Paste",
            "Departmental Records"
        ])
        
        with tab_upload:
            up_col, dl_col = st.columns([2.6, 1.4], gap="small")
            with dl_col:
                st.download_button(
                    label="Download Excel Roster Template",
                    icon=":material/download:",
                    data=generate_blank_excel_template(),
                    file_name="student_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    help="Download standard Excel template."
                )
            with up_col:
                uploaded_file = st.file_uploader(
                    "Upload Candidate Master List (.xlsx, .xls, .docx, .csv, .htm)",
                    type=["xlsx", "xls", "csv", "docx", "htm", "html"],
                    label_visibility="collapsed"
                )

            if uploaded_file is not None:
                if st.session_state.get("last_uploaded_name") != uploaded_file.name:
                    try:
                        n_df, d_reg, d_name = extract_file_to_dataframe(uploaded_file, uploaded_file.name)
                        st.session_state["active_dataframe"] = n_df
                        st.session_state["detected_reg_col"] = d_reg
                        st.session_state["detected_name_col"] = d_name
                        st.session_state["last_uploaded_name"] = uploaded_file.name
                        st.session_state["file_source_name"] = uploaded_file.name
                        st.session_state["students_list"] = parse_dataframe(n_df, d_reg, d_name)
                    except Exception as e:
                        st.error(f"Error reading uploaded file: {e}")

        with tab_paste:
            paste_input = st.text_area(
                "Paste Candidate Rows (REG NO, NAME):",
                height=110,
                placeholder="COE24ND0001\tABDULLAHI MUSA\nCOE24ND0002\tFATIMA YUSUF",
                label_visibility="collapsed"
            )
            if st.button("Import Pasted Records", icon=":material/content_paste:", use_container_width=True):
                if paste_input.strip():
                    st.session_state["active_dataframe"] = None
                    st.session_state["students_list"] = parse_raw_text(paste_input)
                    st.rerun()

        with tab_sample:
            st.caption("Select an existing departmental scoresheet or candidate register:")
            s_col1, s_col2 = st.columns([2, 1])
            with s_col1:
                sample_files = []
                sample_files += [f for f in os.listdir(".") if f.endswith((".docx", ".xlsx", ".xls", ".htm", ".html")) and ("TUE" in f or "sample" in f or "Class" in f)]
                if os.path.exists("Student_list"):
                    sample_files += [os.path.join("Student_list", f) for f in os.listdir("Student_list") if f.endswith((".docx", ".xlsx", ".xls", ".htm", ".html")) and not f.startswith("~$")]
                if not sample_files:
                    sample_files = ["sample_students.xlsx"]
                selected_sample = st.selectbox("Available departmental files:", sample_files, label_visibility="collapsed")
            with s_col2:
                if st.button("Load Candidate Register", icon=":material/folder_open:", use_container_width=True):
                    try:
                        n_df, d_reg, d_name = extract_file_to_dataframe(selected_sample, selected_sample)
                        st.session_state["active_dataframe"] = n_df
                        st.session_state["detected_reg_col"] = d_reg
                        st.session_state["detected_name_col"] = d_name
                        st.session_state["file_source_name"] = os.path.basename(selected_sample)
                        st.session_state["students_list"] = parse_dataframe(n_df, d_reg, d_name)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error loading dataset: {e}")

        # Column Mapping Verification (if spreadsheet was loaded)
        active_df = st.session_state.get("active_dataframe")
        if active_df is not None and not active_df.empty:
            cols = list(active_df.columns)
            detected_reg = st.session_state.get("detected_reg_col")
            detected_name = st.session_state.get("detected_name_col")

            reg_idx = cols.index(detected_reg) if detected_reg in cols else 0
            name_idx = cols.index(detected_name) if detected_name in cols else (1 if len(cols) > 1 else 0)

            with st.expander("Column Mapping Verification (Auto-Detected)", expanded=True):
                st.caption("Confirm that the Registration Number and Candidate Name columns match your document:")
                m_c1, m_c2 = st.columns(2)
                with m_c1:
                    chosen_reg = st.selectbox(
                        "Registration Number Column:",
                        options=cols,
                        index=reg_idx,
                        key=f"col_map_reg_{st.session_state.get('file_source_name', 'default')}"
                    )
                with m_c2:
                    chosen_name = st.selectbox(
                        "Student Full Name Column:",
                        options=cols,
                        index=name_idx,
                        key=f"col_map_name_{st.session_state.get('file_source_name', 'default')}"
                    )

            st.session_state["students_list"] = parse_dataframe(active_df, chosen_reg, chosen_name)

        # Status badge & Preview
        st_count = len(st.session_state["students_list"])
        if st_count > 0:
            st.markdown(
                f'<div style="margin-top:10px;padding:8px 12px;background:#ECFDF5;border:1px solid #A7F3D0;border-radius:6px;font-size:13px;color:#065F46;">'
                f'<span class="material-symbols-outlined icon-success" style="font-size:16px;">check_circle</span> <b>Official Register: {st_count} Candidates Enrolled.</b>'
                f'</div>',
                unsafe_allow_html=True
            )
            with st.expander(f"View Candidate Roster ({st_count} Candidates)"):
                df_roster = pd.DataFrame([
                    {"#": s.original_sn or idx+1, "Registration No": s.reg_no, "Student Full Name": s.name}
                    for idx, s in enumerate(st.session_state["students_list"])
                ])
                st.dataframe(df_roster, use_container_width=True, height=180, hide_index=True)
        else:
            st.caption("Upload a candidate register file or select an existing departmental records file above.")


# ==========================================
# RIGHT COLUMN: Venues, Reshuffle & Generate
# ==========================================
with col_right:
    # Card 3: Venues & Daily Capacities
    with st.container(border=True):
        st.markdown(
            '<div class="section-title"><span class="material-symbols-outlined icon-brand">meeting_room</span> Section 3: Examination Venues & Capacities</div>',
            unsafe_allow_html=True
        )

        preset_venues = presets.get("standard_venues", [])
        preset_names = [v["name"] for v in preset_venues]

        venue_tab1, venue_tab2 = st.tabs([
            "Campus Venue Presets",
            "Add Venue Manually"
        ])

        with venue_tab1:
            p_col1, p_col2 = st.columns([3, 1], gap="small")
            with p_col1:
                selected_presets = st.multiselect(
                    "Select allocated campus halls for session:",
                    options=preset_names,
                    default=[],
                    placeholder="Click to pick halls (e.g. Chem Eng Theater, Hall A-MPRE...)",
                    key="multiselect_presets",
                    label_visibility="collapsed"
                )
            with p_col2:
                if st.button("Add Chosen Venues", icon=":material/add:", use_container_width=True, disabled=len(selected_presets) == 0):
                    existing_names = [v["name"].strip().upper() for v in st.session_state["venues_list"]]
                    for p_name in selected_presets:
                        if p_name.strip().upper() not in existing_names:
                            def_cap = next((v["default_capacity"] for v in preset_venues if v["name"] == p_name), 40)
                            v_id = st.session_state["venue_id_counter"]
                            st.session_state["venue_id_counter"] += 1
                            st.session_state["venues_list"].append({"id": v_id, "name": p_name, "capacity": def_cap})
                    st.rerun()

        with venue_tab2:
            st.caption("Specify additional hall or theater assigned by Central Examination Committee:")
            c_c1, c_c2, c_c3 = st.columns([2.8, 1.4, 1.2], gap="small")
            with c_c1:
                custom_h_name = st.text_input(
                    "Venue Name",
                    placeholder="e.g. ETF THEATER, LAB 4, AUDITORIUM",
                    key="custom_v_name_input",
                    label_visibility="collapsed"
                )
            with c_c2:
                custom_h_cap = st.number_input(
                    "Capacity",
                    min_value=1,
                    max_value=2000,
                    value=50,
                    key="custom_v_cap_input",
                    label_visibility="collapsed"
                )
            with c_c3:
                if st.button("Add Venue", icon=":material/add:", key="btn_add_custom", use_container_width=True):
                    if custom_h_name.strip():
                        new_venue_name = custom_h_name.strip().upper()
                        v_id = st.session_state["venue_id_counter"]
                        st.session_state["venue_id_counter"] += 1
                        st.session_state["venues_list"].append({"id": v_id, "name": new_venue_name, "capacity": int(custom_h_cap)})
                        st.rerun()

        st.markdown("---")

        # Active Venues Header with Clear All button
        v_hdr_l, v_hdr_r = st.columns([3, 1])
        with v_hdr_l:
            st.markdown(f"**Allocated Venues for Session ({len(st.session_state['venues_list'])} Venues):**")
        with v_hdr_r:
            if len(st.session_state["venues_list"]) > 0:
                if st.button("Clear All", icon=":material/delete_sweep:", help="Clear all venues to reconfigure"):
                    st.session_state["venues_list"] = []
                    st.rerun()

        # Render each venue as a stable, persistent row
        venues_to_delete = []
        if len(st.session_state["venues_list"]) == 0:
            st.info("No examination venues assigned yet. Pick halls from the directory above or enter custom halls.")
        else:
            for v_idx, v in enumerate(st.session_state["venues_list"]):
                v_id = v["id"]
                r_c1, r_c2, r_c3 = st.columns([3.4, 1.8, 0.6], gap="small")
                with r_c1:
                    new_name = st.text_input(
                        f"Venue #{v_idx+1}",
                        value=v["name"],
                        key=f"v_name_{v_id}",
                        label_visibility="collapsed"
                    )
                    v["name"] = new_name
                with r_c2:
                    new_cap = st.number_input(
                        f"Cap #{v_idx+1}",
                        min_value=1,
                        max_value=2000,
                        value=v["capacity"],
                        key=f"v_cap_{v_id}",
                        label_visibility="collapsed"
                    )
                    v["capacity"] = int(new_cap)
                with r_c3:
                    if st.button("", icon=":material/delete:", key=f"del_{v_id}", help="Remove this venue"):
                        venues_to_delete.append(v_id)

            if venues_to_delete:
                st.session_state["venues_list"] = [v for v in st.session_state["venues_list"] if v["id"] not in venues_to_delete]
                st.rerun()

        # Live Allocation Balance Indicator
        total_students = len(st.session_state["students_list"])
        total_capacity = sum(int(v["capacity"]) for v in st.session_state["venues_list"])
        diff = total_capacity - total_students

        m1, m2, m3 = st.columns(3)
        with m1:
            st.markdown(f'<div class="metric-pill"><div class="metric-pill-label">Total Candidates</div><div class="metric-pill-value">{total_students}</div></div>', unsafe_allow_html=True)
        with m2:
            st.markdown(f'<div class="metric-pill"><div class="metric-pill-label">Allocated Seats</div><div class="metric-pill-value">{total_capacity}</div></div>', unsafe_allow_html=True)
        with m3:
            if total_students == 0:
                status_txt = "Awaiting Candidates"
                bg_color = "#F8FAFC"
                txt_color = "#64748B"
            elif diff == 0:
                status_txt = "Exact Balance"
                bg_color = "#DCFCE7"
                txt_color = "#15803D"
            elif diff > 0:
                status_txt = f"+{diff} Excess Seats"
                bg_color = "#FEF9C3"
                txt_color = "#A16207"
            else:
                status_txt = f"{-diff} Unseated"
                bg_color = "#FEE2E2"
                txt_color = "#B91C1C"

            st.markdown(f'<div class="metric-pill" style="background:{bg_color};border-color:{bg_color};"><div class="metric-pill-label">Seat Balance</div><div class="metric-pill-value" style="color:{txt_color};">{status_txt}</div></div>', unsafe_allow_html=True)

        # Capacity Balance Actions
        if total_students > 0 and len(st.session_state["venues_list"]) > 0:
            b_c1, b_c2 = st.columns(2)
            with b_c1:
                if st.button("Distribute Evenly Across Venues", icon=":material/balance:", use_container_width=True, help="Divide candidate total equally across all designated halls"):
                    n_halls = len(st.session_state["venues_list"])
                    base = total_students // n_halls
                    rem = total_students % n_halls
                    for i, v in enumerate(st.session_state["venues_list"]):
                        val = base + (1 if i < rem else 0)
                        v["capacity"] = val
                        st.session_state[f"v_cap_{v['id']}"] = val
                    st.rerun()

            with b_c2:
                if diff != 0 and len(st.session_state["venues_list"]) > 0:
                    other_sum = sum(int(v["capacity"]) for v in st.session_state["venues_list"][:-1])
                    needed_for_last = max(1, total_students - other_sum)
                    if st.button("Allocate Balance to Last Venue", icon=":material/vertical_align_bottom:", use_container_width=True, help=f"Set the last venue capacity to {needed_for_last} for 100% balance"):
                        last_v = st.session_state["venues_list"][-1]
                        last_v["capacity"] = needed_for_last
                        st.session_state[f"v_cap_{last_v['id']}"] = needed_for_last
                        st.rerun()

    # Card 4: Reshuffle & Generation
    with st.container(border=True):
        st.markdown(
            '<div class="section-title"><span class="material-symbols-outlined icon-brand">verified_user</span> Section 4: Seating Allocation & Materials Dispatch</div>',
            unsafe_allow_html=True
        )

        col_tgl, col_sd = st.columns([1.1, 1.9])
        with col_tgl:
            enable_shuffle = st.toggle("Malpractice Security Reshuffle", value=True, help="Randomizes candidate seating across halls to prevent peer collusion.")
        with col_sd:
            if enable_shuffle:
                seed_c1, seed_c2 = st.columns([2.8, 1.2])
                with seed_c1:
                    custom_seed_input = st.text_input(
                        "Daily Security Seed Code",
                        value=st.session_state["custom_seed"],
                        help="Alphanumeric daily security code used to randomize seating. Re-entering the same code reproduces the identical seating plan for auditing."
                    )
                with seed_c2:
                    st.write("")
                    if st.button("New Code", icon=":material/refresh:", use_container_width=True, help="Generate random daily security code"):
                        new_rnd = f"SEED_{random.randint(1000, 9999)}"
                        st.session_state["custom_seed"] = new_rnd
                        st.rerun()
            else:
                custom_seed_input = None
                st.caption("Candidates will be seated strictly in their original class register order.")

        st.markdown("<div style='margin-top:6px;'></div>", unsafe_allow_html=True)
        in_hall_order_choice = st.selectbox(
            "In-Hall Candidate Arrangement:",
            options=[
                "Serial by Registration Number (Recommended — Fast Roll Call)",
                "Maintain Randomized Seating Order",
                "Serial by Master List Order",
                "Alphabetical by Candidate Full Name"
            ],
            index=0,
            help="Candidates are shuffled across exam halls based on your daily security code. Inside each venue, this arranges how candidate names appear on the register and door notices."
        )
        order_map = {
            "Serial by Registration Number (Recommended — Fast Roll Call)": "reg_no",
            "Maintain Randomized Seating Order": "random",
            "Serial by Master List Order": "original_sn",
            "Alphabetical by Candidate Full Name": "name"
        }
        selected_in_hall_order = order_map[in_hall_order_choice]

        # Generation Button
        ready_to_generate = (total_students > 0 and total_capacity > 0 and len(st.session_state["venues_list"]) > 0)

        btn_label = "Generate Official Examination Registers & Dispatch Materials" if ready_to_generate else "Specify Course, Students and Venues Above to Generate"
        if st.button(btn_label, icon=":material/task_alt:", type="primary", use_container_width=True, disabled=not ready_to_generate):
            venue_pairs = [(v["name"].strip(), int(v["capacity"])) for v in st.session_state["venues_list"] if str(v["name"]).strip()]

            allocations, summary = allocate_to_venues(
                students=st.session_state["students_list"],
                venues=venue_pairs,
                seed=custom_seed_input if enable_shuffle else None,
                reshuffle=enable_shuffle,
                in_hall_order=selected_in_hall_order
            )

            st.session_state["active_allocations"] = allocations
            st.session_state["active_summary"] = summary
            st.session_state["docx_output"] = build_attendance_docx(exam_config, allocations)
            st.session_state["pdf_output"] = build_attendance_pdf(exam_config, allocations)
            st.session_state["door_docx_output"] = build_door_notices_docx(exam_config, allocations)
            st.session_state["door_pdf_output"] = build_door_notices_pdf(exam_config, allocations)
            st.session_state["envelope_docx_output"] = build_envelope_tags_docx(exam_config, allocations)
            st.session_state["envelope_pdf_output"] = build_envelope_tags_pdf(exam_config, allocations)
            st.session_state["last_course_code"] = exam_config.course_code
            st.session_state["last_level"] = exam_config.level
            st.toast("Attendance Registers, Door Notices, and Envelope Dockets Generated!", icon=":material/check_circle:")


# ==========================================
# RESULTS & DOWNLOADS SECTION
# ==========================================
if "active_allocations" in st.session_state:
    st.markdown("---")
    allocations = st.session_state["active_allocations"]
    summary = st.session_state["active_summary"]

    with st.container(border=True):
        st.markdown(
            '<div class="section-title"><span class="material-symbols-outlined icon-brand">inventory_2</span> Generated Examination Materials & Handover Packets</div>',
            unsafe_allow_html=True
        )

        in_hall_label = summary.get('in_hall_order', 'reg_no').replace('_', ' ').title()
        st.markdown(
            f"**Session Docket Summary:** Security Code: `{summary['seed_used']}` • "
            f"In-Hall Order: `{in_hall_label}` • "
            f"Allocated **{summary['allocated_students']}** of **{summary['total_students']}** candidates across **{len(allocations)}** examination venues."
        )

        clean_code = st.session_state.get("last_course_code", "EXAM").replace(" ", "_")
        clean_lvl = st.session_state.get("last_level", "LEVEL").replace(" ", "_")

        # 1-Click Complete Exam Kit ZIP bundle
        import zipfile
        zip_buf = io.BytesIO()
        zip_filename = f"{clean_code}_{clean_lvl}_COMPLETE_EXAM_PACKET.zip"
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zip_file:
            zip_file.writestr(f"Attendance_Registers/{clean_code}_{clean_lvl}_ATTENDANCE.docx", st.session_state["docx_output"])
            zip_file.writestr(f"Attendance_Registers/{clean_code}_{clean_lvl}_ATTENDANCE.pdf", st.session_state["pdf_output"])
            zip_file.writestr(f"Door_Notices/{clean_code}_{clean_lvl}_DOOR_NOTICES.docx", st.session_state["door_docx_output"])
            zip_file.writestr(f"Door_Notices/{clean_code}_{clean_lvl}_DOOR_NOTICES.pdf", st.session_state["door_pdf_output"])
            zip_file.writestr(f"Envelope_Tags/{clean_code}_{clean_lvl}_ENVELOPE_TAGS.docx", st.session_state["envelope_docx_output"])
            zip_file.writestr(f"Envelope_Tags/{clean_code}_{clean_lvl}_ENVELOPE_TAGS.pdf", st.session_state["envelope_pdf_output"])

        st.download_button(
            label=f"Download Complete Examination Kit (.ZIP) — Registers, Door Notices & Envelope Tags",
            icon=":material/folder_zip:",
            data=zip_buf.getvalue(),
            file_name=zip_filename,
            mime="application/zip",
            use_container_width=True
        )

        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

        export_tab1, export_tab2, export_tab3 = st.tabs([
            "1. Official Attendance Register",
            "2. Venue Entrance Notices (Door Slips)",
            "3. Central Exam Office Envelope Dockets"
        ])

        with export_tab1:
            st.caption("Official examination attendance register for candidate sign-in, booklet serial recording, and invigilator endorsement:")
            d_col1, d_col2 = st.columns(2, gap="medium")
            docx_name = f"{clean_code}_{clean_lvl}_ATTENDANCE.docx"
            pdf_name = f"{clean_code}_{clean_lvl}_ATTENDANCE.pdf"

            with d_col1:
                st.download_button(
                    label=f"Download Attendance Word File ({docx_name})",
                    icon=":material/description:",
                    data=st.session_state["docx_output"],
                    file_name=docx_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    type="primary"
                )

            with d_col2:
                st.download_button(
                    label=f"Download Attendance PDF ({pdf_name})",
                    icon=":material/picture_as_pdf:",
                    data=st.session_state["pdf_output"],
                    file_name=pdf_name,
                    mime="application/pdf",
                    use_container_width=True
                )

        with export_tab2:
            st.caption("High-visibility entrance rosters to affix to venue doors prior to examination commencement:")
            door_col1, door_col2 = st.columns(2, gap="medium")
            door_docx_name = f"{clean_code}_{clean_lvl}_DOOR_NOTICES.docx"
            door_pdf_name = f"{clean_code}_{clean_lvl}_DOOR_NOTICES.pdf"

            with door_col1:
                st.download_button(
                    label=f"Download Door Notices Word ({door_docx_name})",
                    icon=":material/description:",
                    data=st.session_state["door_docx_output"],
                    file_name=door_docx_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    type="primary"
                )

            with door_col2:
                st.download_button(
                    label=f"Download Door Notices PDF ({door_pdf_name})",
                    icon=":material/picture_as_pdf:",
                    data=st.session_state["door_pdf_output"],
                    file_name=door_pdf_name,
                    mime="application/pdf",
                    use_container_width=True
                )

        with export_tab3:
            st.caption("Official packet tags to affix to materials envelopes dispatched to Central Exam Office for Chief Invigilator verification and pick-up:")
            env_col1, env_col2 = st.columns(2, gap="medium")
            env_docx_name = f"{clean_code}_{clean_lvl}_ENVELOPE_TAGS.docx"
            env_pdf_name = f"{clean_code}_{clean_lvl}_ENVELOPE_TAGS.pdf"

            with env_col1:
                st.download_button(
                    label=f"Download Envelope Dockets Word ({env_docx_name})",
                    icon=":material/description:",
                    data=st.session_state["envelope_docx_output"],
                    file_name=env_docx_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    type="primary"
                )

            with env_col2:
                st.download_button(
                    label=f"Download Envelope Dockets PDF ({env_pdf_name})",
                    icon=":material/picture_as_pdf:",
                    data=st.session_state["envelope_pdf_output"],
                    file_name=env_pdf_name,
                    mime="application/pdf",
                    use_container_width=True
                )

        st.markdown("---")
        st.markdown('**Live Venue Seating & Attendance Roster Preview:**')

        tabs = st.tabs([f"{a.venue_name} ({len(a.students)} Candidates)" for a in allocations])
        for idx, tab in enumerate(tabs):
            alloc = allocations[idx]
            with tab:
                st.write(f"**Venue:** `{alloc.venue_name}` | **Assigned Count:** `{len(alloc.students)}` Candidates")
                preview_list = []
                for s_i, s in enumerate(alloc.students):
                    preview_list.append({
                        "S/N": s_i + 1,
                        "REG NO": s.reg_no,
                        "CANDIDATE NAME": s.name,
                        "BOOKLET NO": "[ To be filled by candidate ]",
                        "SIGN IN": "[ Signature ]",
                        "SIGN OUT": "[ Signature ]"
                    })
                st.dataframe(pd.DataFrame(preview_list), use_container_width=True, height=260, hide_index=True)


# Institutional Footer
st.markdown("""
<div style="margin-top: 45px; padding: 22px 10px; border-top: 1px solid #E2E8F0; text-align: center; color: #64748B; font-size: 12.5px; line-height: 1.6;">
    <b>KADUNA POLYTECHNIC</b> • Central Examinations Committee<br/>
    Standardized Examination Hall Allocation & Materials Generation System
</div>
""", unsafe_allow_html=True)
