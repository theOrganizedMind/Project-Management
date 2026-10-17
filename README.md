# Project Management & Scheduling Tool

A desktop GUI application for managing demo/construction projects, tracking tasks, 
generating reports, visualizing schedules in a Gantt chart, and viewing active job locations.

Built with **Python + Tkinter** and currently uses **JSON files** for storage (no SQL database required).

## Features

### 1) Project Management Tab
- Create, update, remove, search, and export projects.
- Fields:
  - Project Type
  - Project Address
  - Job Number
  - Project Manager
  - Crew Name
  - Duration (Days)
  - Bid Price
  - Job Cost
  - P & L (auto-calculated: `Bid Price - Job Cost`)
- Editable dropdowns with starter values:
  - Project Type: `Interior Demolition`, `House Demo`, `Building Demo`
  - Project Manager: `John Doe`, `Jane Smith`
  - Crew Name: `Team 1`, `Team 2`
- Suggests the next available numeric job number.
- Disallows duplicate job numbers.

### 2) Track Project Tab
- Assign and edit tasks by job number.
- Fields:
  - Task
  - Description (optional)
  - Job Number
  - Start Date
  - End Date
  - Status
- Status options:
  - `Upcoming`, `Todo`, `In Progress`, `On Hold`, `Change Order`, `Done`
- Grouped task view by **Job Number + Project Address**.
- Includes:
  - Days Onsite (`End Date - Start Date`)
  - Duration (inclusive days per task)
  - Group-level total duration
- CSV/XLSX import, export, and template generation.

### 3) Reports Tab
- Filterable project-level reporting by:
  - Job Number
  - Address
  - Active Only
- Outputs **one row per project** (not one row per task).
- Aggregates task windows to compute project:
  - Start Date (min task start)
  - End Date (max task end)
  - Days Onsite (sum of inclusive task durations)
  - Aggregate status
- Export report to CSV/XLSX.
- Summary panel with totals and status counts.

### 4) Gantt Chart Tab
- Task timeline visualization with details table.
- Controls:
  - Refresh Chart
  - Show Done toggle
  - Status filter
  - Start Date / End Date filter
- Start Date auto-resets to current date on refresh.

### 5) Map Tab
- Displays active job locations.
- Shows unique `(Job Number, Project Address)` entries.
- Includes an active location list and map-style visual panel.

---

## Data Storage (PUBLIC vs PRIVATE)

The app picks data location automatically:

- If `json_files/` exists: use **PRIVATE** mode  
  - File: `json_files/project_management_private.json`
- Otherwise: use **PUBLIC** mode  
  - File: `data/project_management_public.json`

This behavior is implemented in the app and intended for public-repo safety 
while still allowing local private data.

---

## Requirements

See [requirements.txt](./requirements.txt).

Key packages:
- `tkcalendar` (calendar date picker support)
- `pytest` (unit tests)
- `openpyxl` is optional at runtime for `.xlsx` import/export support (auto-detected by the app if installed)

> Note: `tkinter` is required and is usually bundled with standard Python installers.

---

## Setup

### Windows (PowerShell)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### macOS/Linux

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

---

## Run the Application

```bash
python project_management.py
```

---

## Running Tests

Unit tests live in [tests/test_project_management.py](./tests/test_project_management.py).

Run:

```bash
python -m pytest -q
```

---

## Project Structure

```text
tom_project_management/
├── project_management.py
├── requirements.txt
├── README.md
├── tests/
│   └── test_project_management.py
├── json_files/                  # optional private data dir
│   └── project_management_private.json
└── data/                        # public data dir if json_files doesn't exist
    └── project_management_public.json
```

---

## Notes

- The app uses centralized exception handling for user-safe error dialogs and logs technical details through Python logging.
- The current storage is JSON-based by design; this can be migrated to SQL in a future version.
