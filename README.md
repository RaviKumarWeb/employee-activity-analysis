# Employee Activity Analysis

A Python solution to identify users present in employee activity logs but missing from the active employee list — built to handle large Excel files efficiently without crashing memory.

---

## Problem Statement

Given a large Excel file containing:
- **Sheet 1 (Activity_Log)** — employee activity data with columns: `user_id`, `login_time`, `action`, `department`, `session_duration_mins`, `ip_address`
- **Sheet 2 (Active_Employees)** — official active employee list

**Goal:** Identify users who appear in the activity log but are NOT present in the active employee list (ghost users / ex-employees still active in the system).

---

## Solution Highlights

- ✅ Loads large Excel files using **openpyxl read_only streaming** — reads row by row, never loads full file into memory
- ✅ Handles files of **any size** — tested on 18MB / 500,000 rows without memory issues
- ✅ Auto-detects ID column names — works on any Excel structure
- ✅ Generates a detailed **Excel report** of all missing users with their activity summary

---

## Project Structure

```
employee-activity-analysis/
├── data/
│   ├── employee_activity.xlsx     ← input Excel file (two sheets)
│   └── employee_data.csv          ← source data used to generate dataset
├── solution.py                    ← main solution
├── generate_activity_log.py       ← script used to generate test dataset
├── missing_users_report.xlsx      ← sample output report
└── README.md
```

---

## How to Run

**1. Install dependencies**
```bash
pip install pandas openpyxl xlsxwriter
```

**2. Run the solution**
```bash
python solution.py
```

**3. Output**
- Terminal shows chunk-by-chunk progress and final summary
- `missing_users_report.xlsx` is generated with two sheets:
  - `Missing Users Summary` — one row per ghost user with total actions, departments, first and last activity
  - `Full Activity Detail` — complete activity records of all ghost users

---

## How It Works

```
Excel File
├── Activity_Log Sheet     ──→  Stream in chunks of 50,000 rows
│                                      ↓
└── Active_Employees Sheet ──→  Load into memory as a Set
                                       ↓
                            Compare: activity IDs - employee IDs
                                       ↓
                            Ghost Users Found → Generate Report
```

---

## Sample Output

```
============================================================
   EMPLOYEE ACTIVITY ANALYSIS — MISSING USER FINDER
============================================================

File: data/employee_activity.xlsx (18.0 MB)
Sheets found: ['Activity_Log', 'Active_Employees']

  Total active employees: 3,000
  Processing in chunks of 50,000 rows...

  Chunk  10:   500,000 rows processed | Ghost users found: 10

============================================================
  TOTAL ROWS PROCESSED  : 500,000
  ACTIVE EMPLOYEES      : 3,000
  MISSING / GHOST USERS : 10
============================================================
```

---

## Tech Stack

- Python 3.x
- Pandas
- OpenPyXL
- XlsxWriter