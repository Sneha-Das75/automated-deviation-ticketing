import glob
import os

import pandas as pd
import pdfplumber

import settings

REQUIRED_COLUMNS = [
    "Ref_No", "Program_Name", "Course_Type", "Program_Director",
    "Start_Date", "End_Date", "Deviation_Type", "Attendance", "PQI", "FQI",
]


def find_pdfs():
    pattern = os.path.join(settings.INCOMING_FOLDER, "*.pdf")
    return sorted(glob.glob(pattern), key=os.path.getmtime)


def load_report(pdf_path):
    header = None
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                for row in table:
                    if header is None:
                        header = row
                    elif row != header:
                        rows.append(row)

    if header is None:
        raise ValueError(f"No table found in {os.path.basename(pdf_path)}")

    df = pd.DataFrame(rows, columns=header)
    df.columns = (df.columns.str.replace("\n", " ", regex=False)
                  .str.strip().str.replace(" ", "_"))

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Columns missing in the PDF: {missing}")

    df = df.apply(lambda col: col.str.replace("\n", " ", regex=False).str.strip())
    df = df[df["Ref_No"].notna() & (df["Ref_No"] != "")].copy()

    for col in ["Start_Date", "End_Date"]:
        df[col] = pd.to_datetime(df[col], format="%d %b %Y, %I:%M %p", errors="coerce")
    for col in ["Attendance", "PQI", "FQI"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    bad_dates = df[["Start_Date", "End_Date"]].isna().any(axis=1).sum()
    if bad_dates:
        print(f"Warning: {bad_dates} row(s) have a date that could not be read.")

    df["Base_Program"] = df["Program_Name"].str.replace(
        r"\s+\d{1,2}(st|nd|rd|th)\s+[A-Za-z]{3}\s+\d{4}$", "", regex=True
    )
    df["Dedup_Key"] = df["Ref_No"] + "|" + df["Deviation_Type"]
    df["Load_Date"] = pd.Timestamp.today().normalize()
    df["Source_File"] = os.path.basename(pdf_path)
    return df.reset_index(drop=True)


if __name__ == "__main__":
    pdfs = find_pdfs()
    print("PDFs found in incoming:", len(pdfs))
    if pdfs:
        report = load_report(pdfs[-1])
        print(report.shape)
        print(report[["Ref_No", "Base_Program", "Dedup_Key"]].head(3))