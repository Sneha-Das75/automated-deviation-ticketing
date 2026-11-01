import numpy as np
import pandas as pd

import xl


def create_tickets(report, wb):
    ws_raw = wb["Raw_Intake"]
    ws_master = wb["Master_Deviations"]
    ws_tracker = wb["Ticket_Tracker"]
    ws_history = wb["Status_History"]

    raw_df = xl.sheet_to_df(ws_raw)
    master_df = xl.sheet_to_df(ws_master)
    tracker_df = xl.sheet_to_df(ws_tracker)
    mapping_df = xl.sheet_to_df(wb["Assignment_Mapping"])
    config = xl.sheet_to_df(wb["Config"]).set_index("Setting")["Value"]

    source_file = report["Source_File"].iloc[0]
    if source_file in raw_df["Source_File"].values:
        print(f"{source_file} was already loaded. Skipping it.")
        return {"new": 0, "repeat": 0, "skipped": True}

    sla_days = int(config["SLA_Resolve_Days"])
    known_keys = set(master_df["Dedup_Key"].dropna())
    today = pd.Timestamp.today().normalize()
    last_number = max(
        [int(t.split("-")[-1]) for t in tracker_df["Ticket_ID"].dropna().astype(str)],
        default=0,
    )

    new_tickets = 0
    repeats = 0

    for record in report.to_dict("records"):
        row = {k: (None if pd.isna(v) else v) for k, v in record.items()}
        key = row["Dedup_Key"]

        xl.append_row(ws_raw, row)

        if key in known_keys:
            master_row = xl.find_row(ws_master, "Dedup_Key", key)
            count = xl.get_value(ws_master, master_row, "Times_Reported") or 1
            xl.set_value(ws_master, master_row, "Times_Reported", count + 1)
            xl.set_value(ws_master, master_row, "Last_Seen_Date", today)

            ticket_row = xl.find_row(ws_tracker, "Dedup_Key", key)
            if ticket_row is not None:
                xl.set_value(ws_tracker, ticket_row, "Times_Reported", count + 1)
            repeats += 1
            continue

        known_keys.add(key)
        last_number += 1
        ticket_id = f"DEV-{today.year}-{last_number:06d}"

        match = mapping_df[
            (mapping_df["Program_Director"] == row["Program_Director"])
            & (mapping_df["Deviation_Type"] == row["Deviation_Type"])
        ]
        if match.empty:
            assigned_to, email, status = "UNASSIGNED", None, "Open"
        else:
            assigned_to = row["Program_Director"]
            email = match.iloc[0]["Assignee_Email"]
            status = "Assigned"

        due_date = pd.Timestamp(
            np.busday_offset(np.datetime64(today.date()), sla_days, roll="forward")
        )

        master_record = dict(row)
        master_record.update(
            {"First_Seen_Date": today, "Last_Seen_Date": today, "Times_Reported": 1}
        )
        xl.append_row(ws_master, master_record)

        ticket = dict(row)
        ticket.update({
            "Ticket_ID": ticket_id, "Created_Date": today,
            "Assigned_To": assigned_to, "Assignee_Email": email,
            "Status": status, "Due_Date": due_date, "Times_Reported": 1,
        })
        xl.append_row(ws_tracker, ticket)

        xl.append_row(ws_history, {
            "Ticket_ID": ticket_id, "Old_Status": None, "New_Status": status,
            "Changed_By": "System", "Changed_On": today, "Comment": "Ticket created",
        })
        new_tickets += 1

    return {"new": new_tickets, "repeat": repeats, "skipped": False}


if __name__ == "__main__":
    import intake

    report = intake.load_report(intake.find_pdfs()[-1])
    wb = xl.open_workbook()
    print(create_tickets(report, wb))
    print("Test only: nothing was saved.")