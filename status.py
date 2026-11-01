import pandas as pd

import xl

DATE_FOR_STATUS = {"Resolved": "Resolved_Date", "Closed": "Closed_Date"}


def process_status_changes(wb):
    ws_tracker = wb["Ticket_Tracker"]
    ws_history = wb["Status_History"]

    tracker_df = xl.sheet_to_df(ws_tracker)
    history_df = xl.sheet_to_df(ws_history)

    if tracker_df.empty:
        return {"changed": 0, "dates_filled": 0}

    last_known = history_df.groupby("Ticket_ID")["New_Status"].last()
    current = tracker_df.set_index("Ticket_ID")["Status"]
    assignee = tracker_df.set_index("Ticket_ID")["Assigned_To"]

    compare = pd.DataFrame({"Old_Status": last_known, "New_Status": current})
    changed = compare[compare["Old_Status"] != compare["New_Status"]]
    changed = changed[changed["New_Status"].notna()]

    now = pd.Timestamp.now().floor("s")
    for ticket_id, row in changed.iterrows():
        old = None if pd.isna(row["Old_Status"]) else row["Old_Status"]
        xl.append_row(ws_history, {
            "Ticket_ID": ticket_id,
            "Old_Status": old,
            "New_Status": row["New_Status"],
            "Changed_By": assignee[ticket_id],
            "Changed_On": now,
            "Comment": "Status changed in tracker",
        })

    history_df = xl.sheet_to_df(ws_history)
    filled = 0
    for status, date_column in DATE_FOR_STATUS.items():
        when = (history_df[history_df["New_Status"] == status]
                .groupby("Ticket_ID")["Changed_On"].last())
        for row in ws_tracker.iter_rows(min_row=2):
            ticket_id = xl.get_value(ws_tracker, row, "Ticket_ID")
            if (xl.get_value(ws_tracker, row, "Status") == status
                    and xl.get_value(ws_tracker, row, date_column) is None
                    and ticket_id in when.index):
                xl.set_value(ws_tracker, row, date_column,
                             pd.Timestamp(when[ticket_id]).to_pydatetime())
                filled += 1

    return {"changed": len(changed), "dates_filled": filled}


if __name__ == "__main__":
    wb = xl.open_workbook()
    print("Before any change:", process_status_changes(wb))

    tracker = wb["Ticket_Tracker"]
    row = xl.find_row(tracker, "Ticket_ID", "DEV-2026-000003")
    xl.set_value(tracker, row, "Status", "Resolved")
    print("After pretending ticket 3 was resolved:", process_status_changes(wb))
    print("Test only: nothing was saved.")