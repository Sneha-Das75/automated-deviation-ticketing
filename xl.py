from openpyxl import load_workbook

import settings
import pandas as pd

CLOSE_EXCEL_MESSAGE = (
    "Cannot use the tracker file. Please close Deviation_Tracker.xlsx "
    "in Excel, then run again."
)

def sheet_to_df(sheet):
    data = list(sheet.values)
    return pd.DataFrame(data[1:], columns=data[0])

def open_workbook():
    try:
        return load_workbook(settings.WORKBOOK_FILE)
    except PermissionError:
        raise SystemExit(CLOSE_EXCEL_MESSAGE)


def save_workbook(wb):
    try:
        wb.save(settings.WORKBOOK_FILE)
    except PermissionError:
        raise SystemExit(CLOSE_EXCEL_MESSAGE)


def headers(sheet):
    return [cell.value for cell in sheet[1]]


def append_row(sheet, values):
    sheet.append([values.get(h) for h in headers(sheet)])


def find_row(sheet, column_name, value):
    index = headers(sheet).index(column_name)
    for row in sheet.iter_rows(min_row=2):
        if row[index].value == value:
            return row
    return None


def get_value(sheet, row, column_name):
    return row[headers(sheet).index(column_name)].value


def set_value(sheet, row, column_name, value):
    row[headers(sheet).index(column_name)].value = value


if __name__ == "__main__":
    wb = open_workbook()
    print(wb.sheetnames)
    tracker = wb["Ticket_Tracker"]
    print("Tickets:", tracker.max_row - 1)
    first = find_row(tracker, "Ticket_ID", "DEV-2026-000001")
    print("Status of DEV-2026-000001:", get_value(tracker, first, "Status"))