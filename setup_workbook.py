import os

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import settings

SOURCE_COLUMNS = [
    "Ref_No", "Program_Name", "Course_Type", "Program_Director",
    "Start_Date", "End_Date", "Deviation_Type", "Attendance",
    "PQI", "FQI", "Base_Program", "Dedup_Key", "Load_Date", "Source_File",
]

TABS = {
    "Raw_Intake": SOURCE_COLUMNS,
    "Master_Deviations": SOURCE_COLUMNS + [
        "First_Seen_Date", "Last_Seen_Date", "Times_Reported"],
    "Ticket_Tracker": [
        "Ticket_ID", "Dedup_Key", "Created_Date", "Ref_No", "Program_Name",
        "Base_Program", "Course_Type", "Program_Director", "Start_Date",
        "End_Date", "Deviation_Type", "Attendance", "PQI", "FQI",
        "Assigned_To", "Assignee_Email", "Status", "Due_Date", "Action_Taken",
        "Remarks", "Corrected_Date", "Resolved_Date", "Closed_Date",
        "Validated_By", "Times_Reported"],
    "Status_History": [
        "Ticket_ID", "Old_Status", "New_Status", "Changed_By",
        "Changed_On", "Comment"],
    "Notification_Log": [
        "Ticket_ID", "Notification_Type", "To", "Subject",
        "Sent_On", "Mode", "Result"],
}

DIRECTORS = ["Abhinav", "Rohit", "Neha", "Sanjay", "Priya", "Karan", "Meera"]
DEVIATION_TYPES = ["Attendance", "PQI", "FQI"]


def build_mapping():
    rows = []
    for director in DIRECTORS:
        for deviation_type in DEVIATION_TYPES:
            rows.append({
                "Program_Director": director,
                "Deviation_Type": deviation_type,
                "Assignee_Email": director.lower() + "@example.com",
                "Escalation_Email": "manager@example.com",
                "Validator_Email": "quality@example.com",
            })
    return pd.DataFrame(rows)


def main():
    if os.path.exists(settings.WORKBOOK_FILE):
        print("Deviation_Tracker.xlsx already exists, so nothing was changed.")
        print("To start fresh, rename or delete it first, then run this again.")
        return

    config = pd.DataFrame({
        "Setting": ["SLA_Resolve_Days", "Reminder_Day",
                    "Escalation_Day", "Validation_Days"],
        "Value": [3, 2, 4, 2],
    })

    with pd.ExcelWriter(settings.WORKBOOK_FILE, engine="openpyxl") as writer:
        for tab_name, columns in TABS.items():
            pd.DataFrame(columns=columns).to_excel(
                writer, sheet_name=tab_name, index=False)
        build_mapping().to_excel(writer, sheet_name="Assignment_Mapping", index=False)
        config.to_excel(writer, sheet_name="Config", index=False)

    wb = load_workbook(settings.WORKBOOK_FILE)

    tracker = wb["Ticket_Tracker"]
    header_names = [cell.value for cell in tracker[1]]
    status_col = get_column_letter(header_names.index("Status") + 1)
    dropdown = DataValidation(
        type="list",
        formula1='"Open,Assigned,In Progress,Resolved,Closed"',
        allow_blank=True,
    )
    tracker.add_data_validation(dropdown)
    dropdown.add(f"{status_col}2:{status_col}1000")

    for sheet in wb.worksheets:
        sheet.freeze_panes = "A2"

    wb.save(settings.WORKBOOK_FILE)
    print("Created a fresh", settings.WORKBOOK_FILE)
    print("Tabs:", wb.sheetnames)


if __name__ == "__main__":
    main()