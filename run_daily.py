import os
import shutil

import pandas as pd

import intake
import notify
import settings
import status
import tickets
import xl
import export_dashboard
import intake
import mail_intake
import notify


def archive_pdf(pdf_path):
    os.makedirs(settings.PROCESSED_FOLDER, exist_ok=True)
    name = os.path.basename(pdf_path)
    target = os.path.join(settings.PROCESSED_FOLDER, name)
    if os.path.exists(target):
        stamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        target = os.path.join(settings.PROCESSED_FOLDER, f"{stamp}_{name}")
    shutil.move(pdf_path, target)

def pick_up_mail():
    mode = settings.PICKUP_MODE
    if mode == "FOLDER":
        return
    if mode == "IMAP":
        try:
            saved = mail_intake.fetch_new_pdfs()
        except Exception as error:
            print(f"Could not read the mailbox: {error}")
            return
        print(f"Mailbox: {len(saved)} new PDF(s) saved to incoming")
        return
    print(f"Unknown PICKUP_MODE '{mode}'. Skipping the mailbox step.")


def main():
    print(f"Daily run started. Pickup mode: {settings.PICKUP_MODE} | "
          f"Send mode: {settings.SEND_MODE}")
    wb = xl.open_workbook()

    loaded = []
    total_new = 0
    total_repeat = 0

    pick_up_mail()
    pdfs = intake.find_pdfs()

    pdfs = intake.find_pdfs()
    if not pdfs:
        print("No new PDF in the incoming folder. Continuing with the other checks.")

    for pdf_path in pdfs:
        name = os.path.basename(pdf_path)
        try:
            report = intake.load_report(pdf_path)
        except Exception as error:
            print(f"Could not read {name}: {error}")
            continue

        result = tickets.create_tickets(report, wb)
        print(f"{name}: {result['new']} new tickets, "
              f"{result['repeat']} repeated deviations")
        total_new += result["new"]
        total_repeat += result["repeat"]
        loaded.append(pdf_path)

    changes = status.process_status_changes(wb)
    print(f"Status changes logged: {changes['changed']}, "
          f"dates filled: {changes['dates_filled']}")

    xl.save_workbook(wb)

    new_mail = notify.send_new_ticket_emails(wb)
    reminders = notify.send_reminders(wb)
    escalations = notify.send_escalations(wb)
    resolutions = notify.send_resolution_notices(wb)

    xl.save_workbook(wb)

    try:
        dashboard = export_dashboard.export_dashboard(wb)
        dashboard_message = f"{dashboard['rows']} tickets exported"
    except Exception as error:
        dashboard_message = f"export failed: {error}"

    for pdf_path in loaded:
        archive_pdf(pdf_path)

    print("---------- Summary ----------")
    print(f"New tickets: {total_new} | Repeated deviations: {total_repeat}")
    print(f"Status changes: {changes['changed']}")
    print(f"New-ticket emails: {new_mail}")
    print(f"Reminders: {reminders}")
    print(f"Escalations: {escalations}")
    print(f"Resolution notices: {resolutions}")
    print(f"Dashboard data: {dashboard_message}")
    print("Daily run finished.")


if __name__ == "__main__":
    main()