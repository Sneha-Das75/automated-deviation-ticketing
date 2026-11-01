import os
import random

import numpy as np
import pandas as pd

import settings
from export_dashboard import COLUMNS, OUTPUT_FOLDER, age_bucket, working_days

random.seed(7)

PROGRAMS = {
    "Abhinav": ["Advance Excel", "Excel Dashboards"],
    "Rohit": ["Power BI Fundamentals", "Statistics for Analytics"],
    "Neha": ["Python for Data Analysis", "Advance Excel Formulas"],
    "Sanjay": ["SQL Essentials", "Machine Learning Intro"],
    "Priya": ["Tableau Dashboards"],
    "Karan": ["Business Communication"],
    "Meera": ["Project Management Basics"],
}
COURSE_TYPES = ["Request", "Open", "In-house"]
DEVIATION_TYPES = ["Attendance", "PQI", "FQI"]


def next_business_day(day, offset):
    return pd.Timestamp(np.busday_offset(np.datetime64(day.date()), offset, roll="forward"))


def make_rows():
    today = pd.Timestamp.today().normalize()
    days = pd.bdate_range(today - pd.Timedelta(days=42), today - pd.Timedelta(days=1))
    rows = []
    number = 0

    for day in days:
        for _ in range(random.choice([0, 1, 1, 2, 2, 3])):
            number += 1
            director = random.choice(list(PROGRAMS))
            due = next_business_day(day, 3)
            age = int(np.busday_count(day.date(), today.date()))

            status = None
            resolved = None
            closed = None
            if random.random() < min(0.85, 0.05 + 0.05 * age):
                candidate = next_business_day(day, random.choice([0, 1, 1, 2, 2, 3, 4, 5, 7]))
                if candidate <= today:
                    resolved = candidate
                    closing = next_business_day(resolved, random.choice([0, 1, 2]))
                    if closing <= today:
                        closed = closing
                        status = "Closed"
                    else:
                        status = "Resolved"
            if status is None:
                status = random.choices(["Assigned", "In Progress"], weights=[6, 4])[0]

            is_open = int(status not in ("Resolved", "Closed"))
            open_age = age if is_open else None
            rows.append({
                "Ticket_ID": f"DEV-{day.year}-{number:06d}",
                "Ref_No": f"ABC{2000 + number}",
                "Base_Program": random.choice(PROGRAMS[director]),
                "Course_Type": random.choice(COURSE_TYPES),
                "Program_Director": director,
                "Assigned_To": director,
                "Deviation_Type": random.choices(DEVIATION_TYPES, weights=[5, 3, 2])[0],
                "Created_Date": day,
                "Created_Month": day.strftime("%Y-%m"),
                "Due_Date": due,
                "Status": status,
                "Is_Open": is_open,
                "Is_Overdue": int(is_open == 1 and today > due),
                "Age_Working_Days": open_age,
                "Age_Bucket": age_bucket(open_age),
                "Resolved_Date": resolved,
                "Closed_Date": closed,
                "Resolution_Working_Days": working_days(day, resolved),
                "Times_Reported": random.choices([1, 2], weights=[9, 1])[0],
            })
    return pd.DataFrame(rows)


def main():
    df = make_rows()
    df["Repeat_Error_Count"] = df.groupby(
        ["Program_Director", "Deviation_Type"])["Ticket_ID"].transform("count")
    df["Is_Repeat_Error"] = (df["Repeat_Error_Count"] > 1).astype(int)
    df = df[COLUMNS]

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    path = os.path.join(OUTPUT_FOLDER, "practice_dashboard_tickets.csv")
    df.to_csv(path, index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")
    print(f"Saved {len(df)} practice tickets to {path}")
    print(df["Status"].value_counts())


if __name__ == "__main__":
    main()
