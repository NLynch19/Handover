# ⚡ MOC Electrical Task Manager

A full-featured Streamlit app for managing and reporting MOC (Management of Change) tasks across departments and sites. Designed for real-world handover, progress tracking, and automated reporting.

---

## 🚀 Features

### Entry Mode
- Add new MOC tasks with full metadata
- Update, search, delete, and clear tasks
- Save tasks to Excel for offline use
- Download filtered task snapshots

### Report Mode
- Upload Excel task files
- Filter by department, site, status, and date
- Visualize task progress with charts
- Generate polished Word summaries for handover

---

## 📦 Technologies Used

- `streamlit` – interactive UI
- `pandas` – data manipulation
- `openpyxl` – Excel export
- `python-docx` – Word report generation
- `matplotlib` – progress visualization

---

## 🛠️ How to Run Locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py

