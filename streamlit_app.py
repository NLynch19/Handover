import streamlit as st
import pandas as pd
from datetime import datetime

from form_manager import get_next_id, create_task_entry, track_changes
from excel_manager import load_excel, save_to_excel
from ui_components import show_unsaved_prompt

EXCEL_PATH = "MOC_Tasks.xlsx"
COLUMNS = [
    "ID No", "AREA", "Site", "MOC No", "Assigned Dept", "Assigned Contractor",
    "Project Number", "Project Name", "Project Title", "Project Manager",
    "MOC Coordinator", "Brief Description", "Deliverables", "Deliverables Location",
    "Target Finish", "Progress", "Condition", "Action Holder", "STATUS", "Last Update"
]

st.sidebar.header("📂 Upload Existing Data")
uploaded_file = st.sidebar.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file:
    st.session_state.task_df = pd.read_excel(uploaded_file)
    st.success("✅ Excel file loaded successfully.")
    st.dataframe(st.session_state.task_df.head(10))  # Preview first 10 rows
else:
    st.session_state.task_df = load_excel(EXCEL_PATH, COLUMNS)

st.set_page_config(page_title="MOC Task Manager", layout="wide")
st.title("⚡ MOC Electrical Task Manager")

# Load Excel data
if "task_df" not in st.session_state:
    st.session_state.task_df = load_excel(EXCEL_PATH, COLUMNS)

if "form_modified" not in st.session_state:
    st.session_state.form_modified = False

next_id = get_next_id(st.session_state.task_df)

# Entry Mode
st.header("📥 Entry Mode")
with st.form("moc_entry_form"):
    st.markdown(f"**Next Task ID:** `{next_id}`")
    col1, col2, col3 = st.columns(3)
    area = col1.selectbox("AREA", ["Water", "South", "North", "Other"])
    site = col2.text_input("Site")
    moc_no = col3.text_input("MOC No")

    col4, col5, col6 = st.columns(3)
    assigned_dept = col4.selectbox("Assigned Dept", ["Eng", "Ops", "QA", "Other"])
    contractor = col5.text_input("Assigned Contractor / Engineer")
    project_number = col6.text_input("Project Number")

    col7, col8, col9 = st.columns(3)
    project_name = col7.text_input("Project Name")
    project_title = col8.text_input("Project / MOC Title")
    project_manager = col9.text_input("Project Manager / Engineer")

    moc_coordinator = st.text_input("MOC Coordinator / Planner")
    brief_description = st.text_area("Brief Description")

    col10, col11 = st.columns(2)
    deliverables = col10.text_area("Deliverables & Updates")
    deliverables_location = col11.text_input("Deliverables Location")

    col12, col13, col14, col15 = st.columns(4)
    target_finish = col12.date_input("Target Finish")
    progress = col13.text_input("Progress")
    condition = col14.selectbox("Condition", ["Open", "Closed", "In Progress"])
    action_holder = col15.text_input("Action Holder")

    status = st.text_input("STATUS")

    submitted = st.form_submit_button("➕ Add Task")

st.header("🗑️ Delete Task")

# Choose a task to delete by MOC No
if not st.session_state.task_df.empty:
    moc_options = st.session_state.task_df["MOC No"].dropna().unique().tolist()
    selected_moc = st.selectbox("Select MOC No to delete", moc_options)

    if st.button("❌ Delete Selected Task"):
        before_count = len(st.session_state.task_df)
        st.session_state.task_df = st.session_state.task_df[st.session_state.task_df["MOC No"] != selected_moc]
        after_count = len(st.session_state.task_df)

        save_to_excel(st.session_state.task_df, EXCEL_PATH)
        st.success(f"✅ Deleted task '{selected_moc}' ({before_count - after_count} row removed)")
else:
    st.info("No tasks available to delete.")


if submitted:
    form_data = {
        "area": area,
        "site": site,
        "moc_no": moc_no,
        "assigned_dept": assigned_dept,
        "contractor": contractor,
        "project_number": project_number,
        "project_name": project_name,
        "project_title": project_title,
        "project_manager": project_manager,
        "moc_coordinator": moc_coordinator,
        "brief_description": brief_description,
        "deliverables": deliverables,
        "deliverables_location": deliverables_location,
        "target_finish": target_finish,
        "progress": progress,
        "condition": condition,
        "action_holder": action_holder,
        "status": status
    }

    new_task = create_task_entry(form_data, next_id)
    st.session_state.task_df = pd.concat(
        [st.session_state.task_df, pd.DataFrame([new_task])],
        ignore_index=True
    )
    st.session_state.form_modified = True
    st.success(f"✅ Task '{moc_no}' added.")

# Save prompt
if st.session_state.form_modified:
    if show_unsaved_prompt():
        save_to_excel(st.session_state.task_df, EXCEL_PATH)
        st.session_state.form_modified = False
        st.success("✅ Changes saved to Excel.")
