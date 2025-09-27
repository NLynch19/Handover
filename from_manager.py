from datetime import datetime

def get_next_id(df):
    return int(df["ID No"].max()) + 1 if not df.empty else 1

def track_changes(current, previous):
    return any(current.get(k) != previous.get(k) for k in current)

def create_task_entry(form_data, next_id):
    return {
        "ID No": next_id,
        "AREA": form_data["area"],
        "Site": form_data["site"],
        "MOC No": form_data["moc_no"],
        "Assigned Dept": form_data["assigned_dept"],
        "Assigned Contractor": form_data["contractor"],
        "Project Number": form_data["project_number"],
        "Project Name": form_data["project_name"],
        "Project Title": form_data["project_title"],
        "Project Manager": form_data["project_manager"],
        "MOC Coordinator": form_data["moc_coordinator"],
        "Brief Description": form_data["brief_description"],
        "Deliverables": form_data["deliverables"],
        "Deliverables Location": form_data["deliverables_location"],
        "Target Finish": form_data["target_finish"].strftime("%Y-%m-%d"),
        "Progress": form_data["progress"],
        "Condition": form_data["condition"],
        "Action Holder": form_data["action_holder"],
        "STATUS": form_data["status"],
        "Last Update": datetime.now().strftime("%Y-%m-%d %H:%M")
    }

