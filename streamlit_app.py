import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib.pyplot as plt

st.set_page_config(page_title="📊 MOC Dashboard", layout="centered")
st.title("📊 MOC Task Manager")

# Mode selector
mode = st.radio("Choose Mode", ["Entry Mode", "Report Mode"])

# ---------------- Entry Mode ----------------
if mode == "Entry Mode":
    if "tasks" not in st.session_state:
        st.session_state["tasks"] = []

    with st.form("task_form"):
        department = st.selectbox("Department", ["Ops", "Eng", "Safety", "HR"])
        status = st.selectbox("Status", ["Complete", "Pending", "In Progress"])
        date = st.date_input("Date")
        description = st.text_area("Task Description")
        submitted = st.form_submit_button("Add Task")

        if submitted:
            st.session_state["tasks"].append({
                "Department": department,
                "Status": status,
                "Date": date,
                "Description": description
            })
            st.success("✅ Task added!")

    df = pd.DataFrame(st.session_state["tasks"])
    if not df.empty:
        st.subheader("📋 Current Tasks")
        st.dataframe(df)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Tasks")
        output.seek(0)

        st.download_button(
            label="📥 Download Excel File",
            data=output,
            file_name="moc_tasks.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ---------------- Report Mode ----------------
elif mode == "Report Mode":
    uploaded_file = st.file_uploader("📁 Upload MOC Excel file", type=["xlsx"])

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Error reading Excel file: {e}")
            st.stop()

        required_cols = {"Department", "Status", "Date"}
        if not required_cols.issubset(df.columns):
            st.error(f"Missing required columns: {required_cols}")
            st.stop()

        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df = df.dropna(subset=["Date"])

        st.sidebar.header("🔍 Filter Options")
        selected_depts = st.sidebar.multiselect("🏢 Department", df["Department"].unique())
        selected_status = st.sidebar.multiselect("📌 Status", df["Status"].unique())
        date_range = st.sidebar.date_input("📅 Date Range", [df["Date"].min(), df["Date"].max()])

        filtered_df = df[
            df["Department"].isin(selected_depts) &
            df["Status"].isin(selected_status) &
            df["Date"].between(date_range[0], date_range[1])
        ]

        st.subheader("📊 Completion by Department")
        completion_rate = filtered_df.groupby("Department")["Status"].apply(lambda x: (x == "Complete").mean())
        st.bar_chart(completion_rate)

        st.subheader("📄 Filtered Task Data")
        st.dataframe(filtered_df)

        st.download_button(
            label="📥 Download Filtered CSV",
            data=filtered_df.to_csv(index=False).encode("utf-8"),
            file_name="filtered_moc_data.csv",
            mime="text/csv"
        )

        def create_word_summary(df):
            doc = Document()
            doc.add_heading("MOC Task Summary", level=0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f"Generated on: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}")
            doc.add_page_break()

            total = len(df)
            completed = (df["Status"] == "Complete").sum()
            pending = (df["Status"] == "Pending").sum()
            rate = round(completed / total * 100, 2) if total else 0

            doc.add_heading("Summary Overview", level=1)
            summary_table = doc.add_table(rows=4, cols=2)
            summary_table.style = "Table Grid"
            summary_table.cell(0, 0).text = "Total Tasks"
            summary_table.cell(0, 1).text = str(total)
            summary_table.cell(1, 0).text = "Completed"
            summary_table.cell(1, 1).text = str(completed)
            summary_table.cell(2, 0).text = "Pending"
            summary_table.cell(2, 1).text = str(pending)
            summary_table.cell(3, 0).text = "Completion Rate (%)"
            summary_table.cell(3, 1).text = str(rate)
            doc.add_paragraph()

            fig, ax = plt.subplots()
            completion_rate.plot(kind="bar", ax=ax, color="skyblue")
            ax.set_ylabel("Completion Rate")
            ax.set_title("Completion by Department")
            plt.tight_layout()

            chart_stream = BytesIO()
            plt.savefig(chart_stream, format="png")
            chart_stream.seek(0)
            doc.add_picture(chart_stream, width=Inches(5.5))
            doc.add_paragraph()

            for dept, group in df.groupby("Department"):
                doc.add_heading(dept, level=2)
                table = doc.add_table(rows=1, cols=len(group.columns))
                table.style = "Table Grid"

                hdr_cells = table.rows[0].cells
                for i, col in enumerate(group.columns):
                    hdr_cells[i].text = col
                    hdr_cells[i].paragraphs[0].runs[0].font.bold = True

                for _, row in group.iterrows():
                    row_cells = table.add_row().cells
                    for i, val in enumerate(row):
                        cell_text = str(val)
                        row_cells[i].text = cell_text
                        if group.columns[i] == "Status" and cell_text.lower() == "pending":
                            run = row_cells[i].paragraphs[0].runs[0]
                            run.font.color.rgb = RGBColor(255, 0, 0)

                doc.add_paragraph()

            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            return buffer

        if st.button("📝 Generate Word Summary"):
            word_buffer = create_word_summary(filtered_df)
            st.download_button(
                label="📄 Download Word Summary",
                data=word_buffer,
                file_name="moc_summary.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
#Full MOC dashboard with Entry + Report Mode


