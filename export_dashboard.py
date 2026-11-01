import os

import numpy as np
import pandas as pd

import settings
import xl

OUTPUT_FOLDER = os.path.join(settings.BASE_DIR, "dashboard_data")

COLUMNS = [
    "Ticket_ID", "Ref_No", "Base_Program", "Course_Type", "Program_Director",
    "Assigned_To", "Deviation_Type", "Created_Date", "Created_Month",
    "Due_Date", "Status", "Is_Open", "Is_Overdue", "Age_Working_Days",
    "Age_Bucket", "Resolved_Date", "Closed_Date", "Resolution_Working_Days",
    "Times_Reported", "Repeat_Error_Count", "Is_Repeat_Error",
]


def working_days(start, end):
    if pd.isna(start) or pd.isna(end):
        return None
    return int(np.busday_count(pd.Timestamp(start).date(), pd.Timestamp(end).date()))


def age_bucket(age):
    if pd.isna(age):
        return None
    if age <= 2:
        return "0-2 days"
    if age <= 5:
        return "3-5 days"
    return "6+ days"


def build_dashboard_table(wb):
    df = xl.sheet_to_df(wb["Ticket_Tracker"])
    if df.empty:
        return pd.DataFrame(columns=COLUMNS)

    today = pd.Timestamp.today().normalize()
    for col in ["Created_Date", "Due_Date", "Resolved_Date", "Closed_Date"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    df["Created_Month"] = df["Created_Date"].dt.strftime("%Y-%m")
    df["Is_Open"] = (~df["Status"].isin(["Resolved", "Closed"])).astype(int)
    df["Is_Overdue"] = ((df["Is_Open"] == 1) & (today > df["Due_Date"])).astype(int)

    df["Age_Working_Days"] = [
        working_days(created, today) if is_open == 1 else None
        for created, is_open in zip(df["Created_Date"], df["Is_Open"])
    ]
    df["Age_Bucket"] = df["Age_Working_Days"].apply(age_bucket)
    df["Resolution_Working_Days"] = [
        working_days(created, resolved)
        for created, resolved in zip(df["Created_Date"], df["Resolved_Date"])
    ]

    df["Repeat_Error_Count"] = df.groupby(
        ["Program_Director", "Deviation_Type"])["Ticket_ID"].transform("count")
    df["Is_Repeat_Error"] = (df["Repeat_Error_Count"] > 1).astype(int)

    return df[COLUMNS]


def export_dashboard(wb):
    table = build_dashboard_table(wb)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    path = os.path.join(OUTPUT_FOLDER, "dashboard_tickets.csv")
    table.to_csv(path, index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")
    return {"rows": len(table), "file": path}


if __name__ == "__main__":
    wb = xl.open_workbook()
    print(export_dashboard(wb))