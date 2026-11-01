# Benchmark: this prototype vs common ticketing practice

Professional ticketing and service-management tools (for example ServiceNow,
Jira Service Management, Freshdesk) follow the same basic ideas. This table
compares them with this prototype, and shows the gaps honestly.

## Comparison

| Practice | Common in ticketing tools | This prototype | Gap / next step |
|---|---|---|---|
| Unique ticket ID | Automatic, unique | Yes (`DEV-2026-000001`) | None |
| Automatic intake | Email-to-ticket, forms, APIs | Reads the Looker PDF from the inbox | Outlook version to test at the office |
| Duplicate detection | Duplicates linked or merged | `Dedup_Key` (Ref No + Deviation Type); repeats raise `Times_Reported` | None |
| Assignment rules | Rules by category, team or workload | By Program Director + Deviation Type (`Assignment_Mapping`) | Backup assignee when someone is absent |
| Lifecycle | New, Assigned, In Progress, Resolved, Closed (often also On Hold, Reopened) | Open, Assigned, In Progress, Resolved, Closed | Add On Hold and Reopened |
| SLA and ageing | SLA by priority, business calendar | One SLA in working days (weekends skipped); ageing columns | SLA by severity; public holidays |
| Notifications | Email, chat, mobile push | Email (Gmail, or Outlook at the office) | Chat or Teams messages |
| Reminders and escalation | Multi-level, automatic | One reminder, one escalation to a manager | Several escalation levels |
| Audit trail | Every change logged with user and time | `Status_History` and `Notification_Log` | Real user names; field-level history |
| Validation and closure | Closure codes, verification step | Validator is notified; `Validated_By`; Closed date | Auto-close after N days |
| Recurring errors | Linked to a "problem record" | `Times_Reported`, `Repeat_Error_Count`, `Is_Repeat_Error` | Link repeated tickets to one root problem |
| Root-cause analysis | Root-cause field, categories, reports | Not built | Add a `Root_Cause` column and chart |
| Dashboards | SLA compliance, backlog, resolution time | Dashboard data file ready; dashboard to be built | See dashboard section |
| Access control | Roles and permissions | Excel file permissions only | Assignees can edit any row |
| Scale | Database-backed, many users | One Excel file, one editor at a time | SharePoint list or database for growth |

## Why Python and Excel

The original design used Power Query, SharePoint and Power Automate. The
available Microsoft 365 account did not include the SharePoint license and
related connectors, so the same ideas were built with free tools:

- Power Query (cleaning) is replaced by `intake.py` (pandas)
- SharePoint list or Excel tracker is replaced by `Deviation_Tracker.xlsx`
- Power Automate (flows) is replaced by `tickets.py`, `status.py`, `notify.py`

The design is portable: if the SharePoint license becomes available, the
tracker becomes a SharePoint list and each Python step becomes a flow, without
changing the process.

## Metrics the dashboard supports

Total and pending deviations, open vs resolved, ageing buckets, overdue
tickets, person-wise and program-wise views, resolution time, monthly trend,
recurring and repeat errors. All come from `dashboard_tickets.csv`.

## Limitations to be aware of

- Excel cannot say who edited a cell, so `Changed_By` assumes the assignee.
- Only one person can edit the tracker at a time.
- The daily run is started manually.
- Public holidays are not counted.

## Suggested next phase

1. Add a `Root_Cause` field (dropdown) and a root-cause chart.
2. SLA by severity and a holiday calendar.
3. On Hold and Reopened statuses.
4. Move the tracker to a SharePoint list when the license allows, with
   Power Automate flows for notifications.
5. Schedule the daily run.