import os
import smtplib
import ssl
from email.message import EmailMessage

import pandas as pd
import numpy as np

import settings
import xl

LOG_COLUMNS = ["Ticket_ID", "Notification_Type", "To", "Subject",
               "Sent_On", "Mode", "Result"]
PREVIEW_FOLDER = os.path.join(settings.BASE_DIR, "email_previews")


def ensure_log_sheet(wb):
    if "Notification_Log" not in wb.sheetnames:
        sheet = wb.create_sheet("Notification_Log")
        sheet.append(LOG_COLUMNS)
        sheet.freeze_panes = "A2"
    return wb["Notification_Log"]


def send_email(to_address, subject, body, label):
    recipient = settings.TEST_RECIPIENT or to_address
    mode = settings.SEND_MODE

    if mode == "SMTP":
        if not settings.GMAIL_ADDRESS or not settings.GMAIL_APP_PASSWORD:
            raise ValueError("GMAIL_ADDRESS or GMAIL_APP_PASSWORD missing in .env")
        message = EmailMessage()
        message["From"] = settings.GMAIL_ADDRESS
        message["To"] = recipient
        message["Subject"] = subject
        message.set_content(body)
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(settings.GMAIL_ADDRESS, settings.GMAIL_APP_PASSWORD)
            server.send_message(message)
        return "Sent"

    if mode == "OUTLOOK":
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = subject
        mail.Body = body
        mail.Send()
        return "Sent"

    if mode == "DRYRUN":
        os.makedirs(PREVIEW_FOLDER, exist_ok=True)
        path = os.path.join(PREVIEW_FOLDER, f"{label}.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write(f"To: {recipient}\nSubject: {subject}\n\n{body}")
        return "Preview saved"

    raise ValueError(f"Unknown SEND_MODE: {mode}")


def already_done(log_df, notification_type):
    done = log_df[
        (log_df["Mode"] == settings.SEND_MODE)
        & (log_df["Result"].isin(["Sent", "Preview saved"]))
        & (log_df["Notification_Type"] == notification_type)
    ]
    return set(done["Ticket_ID"])


def send_batch(wb, tickets, notification_type, build_email, address_column):
    log_sheet = ensure_log_sheet(wb)
    sent = 0
    failed = 0
    for _, ticket in tickets.iterrows():
        subject, body = build_email(ticket)
        label = f"{ticket['Ticket_ID']}_{notification_type.replace(' ', '_')}"
        try:
            result = send_email(ticket[address_column], subject, body, label)
        except Exception as error:
            result = f"FAILED: {error}"
        xl.append_row(log_sheet, {
            "Ticket_ID": ticket["Ticket_ID"],
            "Notification_Type": notification_type,
            "To": settings.TEST_RECIPIENT or ticket[address_column],
            "Subject": subject,
            "Sent_On": pd.Timestamp.now().floor("s"),
            "Mode": settings.SEND_MODE,
            "Result": result,
        })
        if result.startswith("FAILED"):
            failed += 1
        else:
            sent += 1
    return {"sent": sent, "failed": failed}


def build_new_ticket_email(t):
    due = pd.to_datetime(t["Due_Date"]).strftime("%d %b %Y")
    subject = (f"[{t['Ticket_ID']}] Deviation assigned: "
               f"{t['Deviation_Type']} - {t['Program_Name']} (due {due})")
    body = f"""Hello {t['Assigned_To']},

A deviation has been assigned to you.

Ticket ID: {t['Ticket_ID']}
Program: {t['Program_Name']}
Ref No: {t['Ref_No']}
Deviation type: {t['Deviation_Type']}
Attendance: {t['Attendance']}   PQI: {t['PQI']}   FQI: {t['FQI']}
Due date: {due}

What to do:
1. Open the Deviation Tracker: {settings.TRACKER_LOCATION}
2. Find your ticket in the Ticket_Tracker tab.
3. Fill in Action_Taken, Remarks and Corrected_Date, and set Status to
   In Progress or Resolved.

Thank you,
Deviation Tracking System
"""
    return subject, body


def send_new_ticket_emails(wb):
    log_sheet = ensure_log_sheet(wb)
    tracker_df = xl.sheet_to_df(wb["Ticket_Tracker"])
    log_df = xl.sheet_to_df(log_sheet)
    if tracker_df.empty:
        return {"sent": 0, "failed": 0}

    done = already_done(log_df, "New ticket")
    pending = tracker_df[
        ~tracker_df["Ticket_ID"].isin(done) & tracker_df["Assignee_Email"].notna()
    ]
    return send_batch(wb, pending, "New ticket", build_new_ticket_email, "Assignee_Email")

def open_tickets_with_age(wb):
    tracker_df = xl.sheet_to_df(wb["Ticket_Tracker"])
    if tracker_df.empty:
        return tracker_df
    today = pd.Timestamp.today().normalize()
    open_df = tracker_df[~tracker_df["Status"].isin(["Resolved", "Closed"])].copy()
    open_df["Working_Days_Open"] = open_df["Created_Date"].apply(
        lambda d: int(np.busday_count(pd.Timestamp(d).date(), today.date()))
    )
    return open_df


def build_reminder_email(t):
    today = pd.Timestamp.today().normalize()
    due_date = pd.to_datetime(t["Due_Date"])
    due = due_date.strftime("%d %b %Y")
    days_left = int(np.busday_count(today.date(), due_date.date()))

    if days_left > 0:
        timing = f"It is due in {days_left} working day(s), on {due}."
    elif days_left == 0:
        timing = f"It is due today, {due}."
    else:
        timing = f"It is overdue. It was due on {due}."

    subject = (f"[{t['Ticket_ID']}] Reminder: deviation still open - "
               f"{t['Deviation_Type']} - {t['Program_Name']}")
    body = f"""Hello {t['Assigned_To']},

This is a reminder that the following deviation is still open.

Ticket ID: {t['Ticket_ID']}
Program: {t['Program_Name']}
Deviation type: {t['Deviation_Type']}
Current status: {t['Status']}
{timing}

What to do:
1. Open the Deviation Tracker: {settings.TRACKER_LOCATION}
2. Find your ticket in the Ticket_Tracker tab.
3. Update Status, Action_Taken, Remarks and Corrected_Date.

Thank you,
Deviation Tracking System
"""
    return subject, body


def build_escalation_email(t):
    due = pd.to_datetime(t["Due_Date"]).strftime("%d %b %Y")
    action = t["Action_Taken"] if pd.notna(t["Action_Taken"]) else "None recorded"
    remarks = t["Remarks"] if pd.notna(t["Remarks"]) else "None recorded"

    subject = (f"[{t['Ticket_ID']}] ESCALATION: deviation open for "
               f"{t['Working_Days_Open']} working days - "
               f"{t['Deviation_Type']} - {t['Program_Name']}")
    body = f"""Hello,

The following deviation has not been resolved within the expected time
and is being escalated to you.

Ticket ID: {t['Ticket_ID']}
Assigned to: {t['Assigned_To']}
Program: {t['Program_Name']}
Deviation type: {t['Deviation_Type']}
Current status: {t['Status']}
Open for: {t['Working_Days_Open']} working days
Due date: {due}
Action taken so far: {action}
Remarks so far: {remarks}

Please follow up with {t['Assigned_To']} or reassign the ticket.
Deviation Tracker: {settings.TRACKER_LOCATION}

Deviation Tracking System
"""
    return subject, body


def send_reminders(wb):
    log_sheet = ensure_log_sheet(wb)
    config = xl.sheet_to_df(wb["Config"]).set_index("Setting")["Value"]
    reminder_day = int(config["Reminder_Day"])
    escalation_day = int(config["Escalation_Day"])

    open_df = open_tickets_with_age(wb)
    if open_df.empty:
        return {"sent": 0, "failed": 0}

    done = already_done(xl.sheet_to_df(log_sheet), "Reminder")
    pending = open_df[
        (open_df["Working_Days_Open"] >= reminder_day)
        & (open_df["Working_Days_Open"] < escalation_day)
        & ~open_df["Ticket_ID"].isin(done)
        & open_df["Assignee_Email"].notna()
    ]
    return send_batch(wb, pending, "Reminder", build_reminder_email, "Assignee_Email")


def send_escalations(wb):
    log_sheet = ensure_log_sheet(wb)
    config = xl.sheet_to_df(wb["Config"]).set_index("Setting")["Value"]
    escalation_day = int(config["Escalation_Day"])
    mapping_df = xl.sheet_to_df(wb["Assignment_Mapping"])

    open_df = open_tickets_with_age(wb)
    if open_df.empty:
        return {"sent": 0, "failed": 0}

    open_df = open_df.merge(
        mapping_df[["Program_Director", "Deviation_Type", "Escalation_Email"]],
        on=["Program_Director", "Deviation_Type"],
        how="left",
    )
    done = already_done(xl.sheet_to_df(log_sheet), "Escalation")
    pending = open_df[
        (open_df["Working_Days_Open"] >= escalation_day)
        & ~open_df["Ticket_ID"].isin(done)
        & open_df["Escalation_Email"].notna()
    ]
    return send_batch(wb, pending, "Escalation", build_escalation_email, "Escalation_Email")

def build_resolution_email(t):
    if pd.notna(t["Corrected_Date"]):
        corrected = pd.to_datetime(t["Corrected_Date"]).strftime("%d %b %Y")
    else:
        corrected = "Not recorded"
    action = t["Action_Taken"] if pd.notna(t["Action_Taken"]) else "None recorded"
    remarks = t["Remarks"] if pd.notna(t["Remarks"]) else "None recorded"

    subject = (f"[{t['Ticket_ID']}] Resolved - please verify: "
               f"{t['Deviation_Type']} - {t['Program_Name']}")
    body = f"""Hello,

{t['Assigned_To']} has marked the following deviation as Resolved.
Please verify the correction.

Ticket ID: {t['Ticket_ID']}
Program: {t['Program_Name']}
Deviation type: {t['Deviation_Type']}
Resolved by: {t['Assigned_To']}
Corrected date: {corrected}
Action taken: {action}
Remarks: {remarks}

If the correction is satisfactory:
1. Open the Deviation Tracker: {settings.TRACKER_LOCATION}
2. In the Ticket_Tracker tab, set this ticket's Status to Closed
   and enter your name in Validated_By.

If it is not satisfactory, set the Status back to In Progress and add a remark.

Deviation Tracking System
"""
    return subject, body


def send_resolution_notices(wb):
    log_sheet = ensure_log_sheet(wb)
    tracker_df = xl.sheet_to_df(wb["Ticket_Tracker"])
    if tracker_df.empty:
        return {"sent": 0, "failed": 0}
    mapping_df = xl.sheet_to_df(wb["Assignment_Mapping"])

    resolved = tracker_df[tracker_df["Status"] == "Resolved"].merge(
        mapping_df[["Program_Director", "Deviation_Type", "Validator_Email"]],
        on=["Program_Director", "Deviation_Type"],
        how="left",
    )
    done = already_done(xl.sheet_to_df(log_sheet), "Resolved notice")
    pending = resolved[
        ~resolved["Ticket_ID"].isin(done) & resolved["Validator_Email"].notna()
    ]
    return send_batch(wb, pending, "Resolved notice", build_resolution_email, "Validator_Email")


if __name__ == "__main__":
    settings.SEND_MODE = "DRYRUN"
    wb = xl.open_workbook()
    tracker = wb["Ticket_Tracker"]
    row = xl.find_row(tracker, "Ticket_ID", "DEV-2026-000003")
    xl.set_value(tracker, row, "Status", "Resolved")
    xl.set_value(tracker, row, "Action_Taken", "Test action")
    print("Resolution notices:", send_resolution_notices(wb))
    print("Test only: nothing was sent and nothing was saved.")