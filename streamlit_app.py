import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import matplotlib.pyplot as plt

st.set_page_config(page_title="MOC Task Manager", layout="wide")
st.title("⚡ MOC Electrical Task Manager")

# Initialize session state
if "task_df" not in st.session_state:
    st.session_state.task_df = pd.DataFrame(columns=[
        "ID No", "AREA", "Site", "MOC No", "Assigned Dept", "Assigned Contractor",
        "Project Number", "Project Name", "Project Title", "Project Manager",
        "MOC Coordinator", "Brief Description", "Deliverables", "Deliverables Location",
        "Target Finish", "Progress", "Condition", "Action Holder"
    ])

# Entry Mode
st.header("📥 Entry Mode")

with st.form("moc_entry_form"):
    col1, col2, col3 = st.columns(3)
    id_no = col1.text_input("ID No")
    area = col2.selectbox("AREA", ["Water", "South", "North", "Other"])
    site = col3.text_input("Site")

    col4, col5, col6 = st.columns(3)
    moc_no = col4.text_input("MOC No")
    assigned_dept = col5.selectbox("Assigned Dept", ["Eng", "Ops", "QA", "Other"])
    contractor = col6.text_input("Assigned Contractor / Engineer")

    col7, col8, col9 = st.columns(3)
    project_number = col7.text_input("Project Number")
    project_name = col8.text_input("Project Name")
    project_title = col9.text_input("Project / MOC Title")

    col10, col11 = st.columns(2)
    project_manager = col10.text_input("Project Manager / Engineer")
    moc_coordinator = col11.text_input("MOC Coordinator / Planner")

    brief_description = st.text_area("Brief Description")

    col12, col13 = st.columns(2)
    deliverables = col12.text_area("Deliverables & Updates")
    deliverables_location = col13.text_input("Deliverables Location")

    col14, col15, col16, col17 = st.columns(4)
    target_finish = col14.date_input("Target Finish")
    progress = col15.text_input("Progress")
    condition = col16.selectbox("Condition", ["Open", "Closed", "In Progress"])
    action_holder = col17.text_input("Action Holder")

    submitted = st.form_submit_button("➕ Add Task")
    if submitted:
        new_task = {
            "ID No": id_no,
            "AREA": area,
            "Site": site,
            "MOC No": moc_no,
            "Assigned Dept": assigned_dept,
            "Assigned Contractor": contractor,
            "Project Number": project_number,
            "Project Name": project_name,
            "Project Title": project_title,
            "Project Manager": project_manager,
            "MOC Coordinator": moc_coordinator,
            "Brief Description": brief_description,
            "Deliverables": deliverables,
            "Deliverables Location": deliverables_location,
            "Target Finish": target_finish.strftime("%Y-%m-%d"),
            "Progress": progress,
            "Condition": condition,
            "Action Holder": action_holder
        }

        if moc_no in st.session_state.task_df["MOC No"].values:
            st.session_state.task_df.loc[
                st.session_state.task_df["MOC No"] == moc_no
            ] = new_task
            st.success(f"✅ Task '{moc_no}' updated.")
        else:
            st.session_state.task_df = pd.concat(
                [st.session_state.task_df, pd.DataFrame([new_task])],
                ignore_index=True
            )
            st.success(f"✅ Task '{moc_no}' added.")

# Display and manage tasks
st.subheader("📋 Current Tasks")
st.dataframe(st.session_state.task_df)

search_moc = st.text_input("🔍 Search by MOC No")
if search_moc:
    result = st.session_state.task_df[
        st.session_state.task_df["MOC No"] == search_moc
    ]
    st.write(result if not result.empty else "No match found.")

if st.button("🗑️ Delete Task"):
    st.session_state.task_df = st.session_state.task_df[
        st.session_state.task_df["MOC No"] != search_moc
    ]
    st.success(f"Task '{search_moc}' deleted.")

if st.button("🧹 Clear All Tasks"):
    st.session_state.task_df = st.session_state.task_df.iloc[0:0]
    st.success("All tasks cleared.")

# Excel export
def get_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

excel_data = get_excel_download(st.session_state.task_df)
st.download_button(
    label="📥 Download Excel",
    data=excel_data,
    file_name="MOC_Tasks.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Report Mode
st.header("📊 Report Mode")

uploaded_file = st.file_uploader("📤 Upload MOC Excel File", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ File uploaded and parsed.")

    col1, col2, col3 = st.columns(3)
    filter_dept = col1.selectbox("Assigned Dept", ["All"] + sorted(df["Assigned Dept"].dropna().unique()))
    filter_status = col2.selectbox("Condition", ["All"] + sorted(df["Condition"].dropna().unique()))
    filter_site = col3.selectbox("Site", ["All"] + sorted(df["Site"].dropna().unique()))

    filtered_df = df.copy()
    if filter_dept != "All":
        filtered_df = filtered_df[filtered_df["Assigned Dept"] == filter_dept]
    if filter_status != "All":
        filtered_df = filtered_df[filtered_df["Condition"] == filter_status]
    if filter_site != "All":
        filtered_df = filtered_df[filtered_df["Site"] == filter_site]

    st.dataframe(filtered_df)

    st.subheader("📊 Progress Overview")
    progress_counts = filtered_df["Condition"].value_counts()
    st.bar_chart(progress_counts)

    def generate_word_summary(df, filename="MOC_Report.docx"):
        doc = Document()
        doc.add_heading("MOC Task Summary", level=1)

        section = doc.sections[0]
        header = section.header
        header_paragraph = header.paragraphs[0]
        header_paragraph.text = "MOC Electrical Task Manager"
        header_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        footer = section.footer
        footer_paragraph = footer.paragraphs[0]
        footer_paragraph.text = "Page "
        footer_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        grouped = df.groupby("Assigned Dept")
        for dept, group in grouped:
            doc.add_heading(f"Department: {dept}", level=2)
            for _, row in group.iterrows():
                doc.add_heading(f"MOC No: {row['MOC No']}", level=3)
                doc.add_paragraph(f"Project Title: {row['Project Title']}")
                doc.add_paragraph(f"Site: {row['Site']}")
                doc.add_paragraph(f"Condition: {row['Condition']}")
                doc.add_paragraph(f"Target Finish: {row['Target Finish']}")
                doc.add_paragraph(f"Progress: {row['Progress']}")
                doc.add_paragraph(f"Action Holder: {row['Action Holder']}")
                doc.add_paragraph(f"Brief Description:\n{row['Brief Description']}")
                doc.add_paragraph(f"Deliverables:\n{row['Deliverables']}")
                doc.add_paragraph("-" * 40)

        doc.save(filename)

    if st.button("📝 Generate Word Summary"):
        generate_word_summary(filtered_df)
        with open("MOC_Report.docx", "rb") as f:
            word_data = f.read()
        st.download_button(
            label="📥 Download Word Summary",
            data=word_data,
            file_name="MOC_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
