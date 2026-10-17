from __future__ import annotations

from datetime import date, datetime
import csv
import importlib
import importlib.util
import json
import logging
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import uuid

if importlib.util.find_spec("openpyxl") is not None:
    _openpyxl = importlib.import_module("openpyxl")
    Workbook = _openpyxl.Workbook
    load_workbook = _openpyxl.load_workbook
else:
    Workbook = None
    load_workbook = None

if importlib.util.find_spec("tkcalendar") is not None:
    _tkcalendar = importlib.import_module("tkcalendar")
    DateEntry = _tkcalendar.DateEntry
else:
    DateEntry = None

LOGGER = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)

DATE_FMT = "%Y-%m-%d"
TRACK_STATUSES = ("Upcoming", "Todo", "In Progress", "On Hold", "Change Order", "Done")
PROJECT_TYPE_OPTIONS = ("Interior Demolition", "House Demo", "Building Demo")
PROJECT_MANAGER_OPTIONS = ("John Doe", "Jane Smith")
CREW_OPTIONS = ("Team 1", "Team 2")


def safe_float(value: str, field_name: str) -> float:
    """Convert user text to float and raise a clear ValueError on invalid input."""
    text = (value or "").strip().replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError as exc:
        raise ValueError(f"{field_name} must be a valid number.") from exc


def safe_int(value: str, field_name: str) -> int:
    """Convert user text to integer and raise a clear ValueError on invalid input."""
    text = (value or "").strip()
    if not text:
        return 0
    try:
        return int(text)
    except ValueError as exc:
        raise ValueError(f"{field_name} must be a valid whole number.") from exc


def parse_date(value: str, field_name: str) -> date:
    """Parse a YYYY-MM-DD string into a date with user-friendly validation."""
    text = (value or "").strip()
    if not text:
        raise ValueError(f"{field_name} is required.")
    try:
        return datetime.strptime(text, DATE_FMT).date()
    except ValueError as exc:
        raise ValueError(f"{field_name} must be in YYYY-MM-DD format.") from exc


def format_currency(value: float) -> str:
    """Format numeric value as currency for table presentation."""
    return f"${value:,.2f}"


def calculate_days_onsite(start_date: str, end_date: str) -> int:
    """Calculate days onsite as the difference between end date and start date."""
    try:
        start = parse_date(start_date, "Start Date")
        end = parse_date(end_date, "End Date")
    except ValueError:
        return 0
    return max((end - start).days, 0)


def calculate_duration_days(start_date: str, end_date: str) -> int:
    """Calculate inclusive duration from start date through end date."""
    try:
        start = parse_date(start_date, "Start Date")
        end = parse_date(end_date, "End Date")
    except ValueError:
        return 0
    return max((end - start).days + 1, 0)


class ProjectManagementApp:
    """Tkinter application for project management, scheduling, reporting, gantt, and map views."""

    def __init__(self, root: tk.Tk) -> None:
        """Initialize state, load data, build UI, and render initial views."""
        self.root = root
        self.root.title("Project Management")
        self.root.geometry("1360x800")
        self.root.minsize(1200, 700)

        self.data_directory, self.data_visibility = self._resolve_data_directory()
        self.data_file = os.path.join(
            self.data_directory,
            f"project_management_{self.data_visibility.lower()}.json",
        )

        self.data: dict[str, list[dict[str, object]]] = {"projects": [], "tasks": []}
        self.selected_project_id: str | None = None
        self.selected_task_id: str | None = None
        self.task_row_lookup: dict[str, str] = {}
        self.report_rows: list[dict[str, object]] = []
        self.gantt_rows: list[dict[str, object]] = []

        self.active_only_var = tk.BooleanVar(value=True)
        self.gantt_show_done_var = tk.BooleanVar(value=True)
        self.gantt_status_filter_var = tk.StringVar(value="All Statuses")

        self.load_data()
        self.build_ui()
        self._safe_execute(self.refresh_all_views, title="Startup Error")
        self._safe_execute(self.clear_project_form, title="Startup Error")

    def _resolve_data_directory(self) -> tuple[str, str]:
        """Choose private json_files directory when available; otherwise use public data directory."""
        base_dir = os.path.dirname(os.path.abspath(__file__))
        private_dir = os.path.join(base_dir, "json_files")
        public_dir = os.path.join(base_dir, "data")
        if os.path.isdir(private_dir):
            return private_dir, "PRIVATE"
        os.makedirs(public_dir, exist_ok=True)
        return public_dir, "PUBLIC"

    def _safe_execute(self, action, *, title: str = "Operation Failed") -> None:
        """Run a callback with centralized exception handling and user-safe messaging."""
        try:
            action()
        except ValueError as exc:
            LOGGER.warning("Validation failure during %s: %s", getattr(action, "__name__", "action"), exc)
            messagebox.showwarning("Validation Error", str(exc))
        except (OSError, json.JSONDecodeError) as exc:
            LOGGER.exception("Data access failure during %s: %s", getattr(action, "__name__", "action"), exc)
            messagebox.showerror(title, "A data access error occurred. Please try again.")
        except Exception as exc:
            LOGGER.exception("Unexpected failure during %s: %s", getattr(action, "__name__", "action"), exc)
            messagebox.showerror(title, "An unexpected error occurred. Please try again.")

    def _get_date_value(self, widget) -> str:
        """Return normalized YYYY-MM-DD value from a date widget."""
        if DateEntry is not None and isinstance(widget, DateEntry):
            return widget.get_date().isoformat()
        return widget.get().strip()

    def _set_date_value(self, widget, value: str) -> None:
        """Set a date widget to a YYYY-MM-DD value."""
        if DateEntry is not None and isinstance(widget, DateEntry):
            widget.set_date(parse_date(value, "Date"))
            return
        widget.delete(0, tk.END)
        widget.insert(0, value)

    def _project_status_from_tasks(self, tasks: list[dict[str, object]]) -> str:
        """Calculate an aggregate project status from task statuses."""
        statuses = {str(task.get("status", "")).strip() for task in tasks if str(task.get("status", "")).strip()}
        if not statuses:
            return "Not Scheduled"
        if statuses == {"Done"}:
            return "Done"
        for candidate in ("In Progress", "On Hold", "Change Order", "Upcoming", "Todo"):
            if candidate in statuses:
                return candidate
        return sorted(statuses)[0]

    def _suggest_next_job_number(self) -> str:
        """Return the next available numeric job number suggestion."""
        numeric_values: list[int] = []
        for project in self.data["projects"]:
            job_number = str(project.get("job_number", "")).strip()
            if job_number.isdigit():
                numeric_values.append(int(job_number))
        if not numeric_values:
            return "1000"
        return str(max(numeric_values) + 1)

    def load_data(self) -> None:
        """Load project and task JSON data from the selected repository data file."""
        if not os.path.exists(self.data_file):
            self.save_data()
            return
        with open(self.data_file, "r", encoding="utf-8") as file:
            loaded = json.load(file)
        projects = loaded.get("projects", [])
        tasks = loaded.get("tasks", [])
        if not isinstance(projects, list) or not isinstance(tasks, list):
            raise ValueError("Data file structure is invalid.")
        self.data = {"projects": projects, "tasks": tasks}

    def save_data(self) -> None:
        """Persist in-memory project and task data to JSON."""
        os.makedirs(self.data_directory, exist_ok=True)
        with open(self.data_file, "w", encoding="utf-8") as file:
            json.dump(self.data, file, indent=2)

    def build_ui(self) -> None:
        """Build notebook tabs and all tab-specific controls."""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_projects = ttk.Frame(self.notebook)
        self.tab_tracking = ttk.Frame(self.notebook)
        self.tab_reports = ttk.Frame(self.notebook)
        self.tab_gantt = ttk.Frame(self.notebook)
        self.tab_map = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_projects, text="📕Project Management")
        self.notebook.add(self.tab_tracking, text="📍Track Project")
        self.notebook.add(self.tab_reports, text="📄Reports")
        self.notebook.add(self.tab_gantt, text="📈Gantt Chart")
        self.notebook.add(self.tab_map, text="🗺️Map")

        self.build_project_management_tab()
        self.build_track_project_tab()
        self.build_reports_tab()
        self.build_gantt_tab()
        self.build_map_tab()

    def build_project_management_tab(self) -> None:
        """Build Project Management tab form, actions, and table."""
        form = ttk.LabelFrame(
            self.tab_projects,
            text=f"Project Management ({self.data_visibility} data file: {os.path.basename(self.data_file)})",
        )
        form.grid(row=0, column=0, sticky="ew", padx=10, pady=8)

        labels = (
            "Project Type",
            "Project Address",
            "Job Number",
            "Project Manager",
            "Crew Name",
            "Duration (Days)",
            "Bid Price",
            "Job Cost",
            "P & L",
        )
        self.project_entries: dict[str, tk.Widget] = {}
        combo_defaults = {
            "Project Type": PROJECT_TYPE_OPTIONS,
            "Project Manager": PROJECT_MANAGER_OPTIONS,
            "Crew Name": CREW_OPTIONS,
        }
        for idx, label in enumerate(labels):
            col = (idx % 3) * 2
            row = idx // 3
            ttk.Label(form, text=label).grid(row=row, column=col, sticky="e", padx=6, pady=6)
            if label in combo_defaults:
                entry = ttk.Combobox(form, state="normal", width=33, values=combo_defaults[label])
            else:
                entry = ttk.Entry(form, width=36)
            entry.grid(row=row, column=col + 1, sticky="w", padx=6, pady=6)
            if label == "P & L":
                entry.configure(state="readonly")
            self.project_entries[label] = entry

        self.project_entries["Bid Price"].bind("<KeyRelease>", lambda _: self.update_project_pl_preview())
        self.project_entries["Job Cost"].bind("<KeyRelease>", lambda _: self.update_project_pl_preview())

        buttons = ttk.Frame(form)
        buttons.grid(row=3, column=0, columnspan=6, sticky="w", padx=6, pady=(4, 8))
        ttk.Button(buttons, text="Add Project", command=lambda: self._safe_execute(self.add_project)).pack(side="left", padx=5)
        ttk.Button(buttons, text="Update Project", command=lambda: self._safe_execute(self.update_project)).pack(side="left", padx=5)
        ttk.Button(buttons, text="Remove Project", command=lambda: self._safe_execute(self.remove_project)).pack(side="left", padx=5)
        ttk.Button(buttons, text="Search Project", command=lambda: self._safe_execute(self.search_project)).pack(side="left", padx=5)
        ttk.Button(buttons, text="Clear Results", command=lambda: self._safe_execute(self.clear_project_results)).pack(side="left", padx=5)
        ttk.Button(buttons, text="Export", command=lambda: self._safe_execute(self.export_projects)).pack(side="left", padx=5)

        table_frame = ttk.LabelFrame(self.tab_projects, text="Projects")
        table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)

        columns = (
            "project_type",
            "project_address",
            "job_number",
            "project_manager",
            "crew_name",
            "duration_days",
            "bid_price",
            "job_cost",
            "p_and_l",
        )
        self.project_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)
        headings = (
            "Project Type",
            "Project Address",
            "Job Number",
            "Project Manager",
            "Crew Name",
            "Duration (Days)",
            "Bid Price",
            "Job Cost",
            "P & L",
        )
        for col, text in zip(columns, headings):
            self.project_tree.heading(col, text=text)
            width = 140 if col not in {"project_address", "project_manager"} else 210
            self.project_tree.column(col, width=width, anchor="w")

        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.project_tree.yview)
        self.project_tree.configure(yscrollcommand=scroll.set)
        self.project_tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        self.project_tree.bind("<<TreeviewSelect>>", self.on_project_selected)

        self.tab_projects.rowconfigure(1, weight=1)
        self.tab_projects.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

    def build_track_project_tab(self) -> None:
        """Build Track Project tab controls and grouped task table."""
        form = ttk.LabelFrame(self.tab_tracking, text="Task Assignment")
        form.grid(row=0, column=0, sticky="ew", padx=10, pady=8)

        ttk.Label(form, text="Task").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        self.task_entry = ttk.Entry(form, width=34)
        self.task_entry.grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="Description").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.description_entry = ttk.Entry(form, width=44)
        self.description_entry.grid(row=0, column=3, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="Job Number").grid(row=1, column=0, sticky="e", padx=6, pady=6)
        self.task_job_entry = ttk.Combobox(form, state="normal", width=32)
        self.task_job_entry.grid(row=1, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="Start Date").grid(row=1, column=2, sticky="e", padx=6, pady=6)
        if DateEntry is not None:
            self.start_date_entry = DateEntry(form, width=19, date_pattern="yyyy-mm-dd")
        else:
            self.start_date_entry = ttk.Entry(form, width=22)
            self.start_date_entry.insert(0, date.today().isoformat())
        self.start_date_entry.grid(row=1, column=3, sticky="w", padx=6, pady=6)
        if DateEntry is not None:
            self.start_date_entry.set_date(date.today())

        ttk.Label(form, text="End Date").grid(row=2, column=2, sticky="e", padx=6, pady=6)
        if DateEntry is not None:
            self.end_date_entry = DateEntry(form, width=19, date_pattern="yyyy-mm-dd")
        else:
            self.end_date_entry = ttk.Entry(form, width=22)
            self.end_date_entry.insert(0, date.today().isoformat())
        self.end_date_entry.grid(row=2, column=3, sticky="w", padx=6, pady=6)
        if DateEntry is not None:
            self.end_date_entry.set_date(date.today())

        ttk.Label(form, text="Status").grid(row=2, column=0, sticky="e", padx=6, pady=6)
        self.status_combo = ttk.Combobox(form, state="readonly", width=32, values=TRACK_STATUSES)
        self.status_combo.grid(row=2, column=1, sticky="w", padx=6, pady=6)
        self.status_combo.set("Upcoming")

        buttons = ttk.Frame(form)
        buttons.grid(row=3, column=0, columnspan=4, sticky="w", padx=6, pady=(4, 8))
        ttk.Button(buttons, text="Assign Task", command=lambda: self._safe_execute(self.assign_task)).pack(side="left", padx=5)
        ttk.Button(buttons, text="Mark Selected Completed", command=lambda: self._safe_execute(self.mark_selected_completed)).pack(side="left", padx=5)
        ttk.Button(buttons, text="Refresh", command=lambda: self._safe_execute(self.refresh_task_tree)).pack(side="left", padx=5)
        ttk.Button(buttons, text="Clear", command=lambda: self._safe_execute(self.clear_task_form)).pack(side="left", padx=5)
        ttk.Button(buttons, text="Edit", command=lambda: self._safe_execute(self.edit_selected_task)).pack(side="left", padx=5)
        ttk.Button(buttons, text="Import", command=lambda: self._safe_execute(self.import_tasks)).pack(side="left", padx=5)
        ttk.Button(buttons, text="Export", command=lambda: self._safe_execute(self.export_tasks)).pack(side="left", padx=5)
        ttk.Button(buttons, text="Template", command=lambda: self._safe_execute(self.save_task_template)).pack(side="left", padx=5)

        table_frame = ttk.LabelFrame(self.tab_tracking, text="Tasks Grouped by Job Number and Project Address")
        table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)
        columns = ("task", "job_number", "start_date", "end_date", "days_onsite", "duration", "status")
        self.task_tree = ttk.Treeview(table_frame, columns=columns, show="tree headings", height=15)
        self.task_tree.heading("#0", text="Group")
        self.task_tree.column("#0", width=340, anchor="w")
        headings = ("Task", "Job Number", "Start Date", "End Date", "Days Onsite", "Duration", "Status")
        widths = (220, 150, 130, 130, 100, 90, 130)
        for col, text, width in zip(columns, headings, widths):
            self.task_tree.heading(col, text=text)
            self.task_tree.column(col, width=width, anchor="w")

        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.task_tree.yview)
        self.task_tree.configure(yscrollcommand=scroll.set)
        self.task_tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        self.task_tree.bind("<<TreeviewSelect>>", self.on_task_selected)

        self.tab_tracking.rowconfigure(1, weight=1)
        self.tab_tracking.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

    def build_reports_tab(self) -> None:
        """Build Reports tab filters, actions, results tree, and summary pane."""
        filter_frame = ttk.LabelFrame(self.tab_reports, text="Report Filters")
        filter_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=8)

        ttk.Label(filter_frame, text="Job Number").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        self.report_job_combo = ttk.Combobox(filter_frame, state="normal", width=30)
        self.report_job_combo.grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(filter_frame, text="Address").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.report_address_combo = ttk.Combobox(filter_frame, state="normal", width=50)
        self.report_address_combo.grid(row=0, column=3, sticky="w", padx=6, pady=6)

        ttk.Checkbutton(filter_frame, text="Active Only", variable=self.active_only_var).grid(
            row=0,
            column=4,
            sticky="w",
            padx=6,
            pady=6,
        )

        actions = ttk.Frame(filter_frame)
        actions.grid(row=1, column=0, columnspan=5, sticky="w", padx=6, pady=(2, 8))
        ttk.Button(actions, text="Run Filtered Report", command=lambda: self._safe_execute(self.run_filtered_report)).pack(side="left", padx=5)
        ttk.Button(actions, text="All Active Projects", command=lambda: self._safe_execute(self.show_all_active_projects)).pack(side="left", padx=5)
        ttk.Button(actions, text="Project History", command=lambda: self._safe_execute(self.show_project_history)).pack(side="left", padx=5)
        ttk.Button(actions, text="Export", command=lambda: self._safe_execute(self.export_reports)).pack(side="left", padx=5)

        table_frame = ttk.LabelFrame(self.tab_reports, text="Report Results")
        table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)
        columns = (
            "job_number",
            "project_address",
            "project_type",
            "start_date",
            "end_date",
            "days_onsite",
            "status",
            "bid_price",
            "job_cost",
            "p_and_l",
        )
        self.report_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=14)
        headings = (
            "Job Number",
            "Project Address",
            "Project Type",
            "Start Date",
            "End Date",
            "Days Onsite",
            "Status",
            "Bid Price",
            "Job Cost",
            "P & L",
        )
        widths = (120, 260, 150, 120, 120, 110, 120, 110, 110, 110)
        for col, text, width in zip(columns, headings, widths):
            self.report_tree.heading(col, text=text)
            self.report_tree.column(col, width=width, anchor="w")

        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.report_tree.yview)
        self.report_tree.configure(yscrollcommand=scroll.set)
        self.report_tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        summary_frame = ttk.LabelFrame(self.tab_reports, text="Summary View")
        summary_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 8))
        self.summary_text = scrolledtext.ScrolledText(summary_frame, height=9, wrap=tk.WORD)
        self.summary_text.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.summary_text.configure(state="disabled")

        self.tab_reports.rowconfigure(1, weight=1)
        self.tab_reports.rowconfigure(2, weight=1)
        self.tab_reports.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        summary_frame.rowconfigure(0, weight=1)
        summary_frame.columnconfigure(0, weight=1)

    def build_gantt_tab(self) -> None:
        """Build Gantt tab controls, chart canvas, and chart details tree."""
        controls = ttk.LabelFrame(self.tab_gantt, text="Chart Controls")
        controls.grid(row=0, column=0, sticky="ew", padx=10, pady=8)
        ttk.Checkbutton(controls, text="Show Done", variable=self.gantt_show_done_var).grid(
            row=0,
            column=0,
            sticky="w",
            padx=6,
            pady=6,
        )
        ttk.Label(controls, text="Status").grid(row=0, column=1, sticky="e", padx=6, pady=6)
        self.gantt_status_combo = ttk.Combobox(
            controls,
            state="readonly",
            width=20,
            textvariable=self.gantt_status_filter_var,
            values=("All Statuses",) + TRACK_STATUSES,
        )
        self.gantt_status_combo.grid(row=0, column=2, sticky="w", padx=6, pady=6)
        self.gantt_status_combo.set("All Statuses")

        ttk.Label(controls, text="Start Date").grid(row=0, column=3, sticky="e", padx=6, pady=6)
        if DateEntry is not None:
            self.gantt_start_filter = DateEntry(controls, width=14, date_pattern="yyyy-mm-dd")
            self.gantt_start_filter.set_date(date.today())
        else:
            self.gantt_start_filter = ttk.Entry(controls, width=17)
            self.gantt_start_filter.insert(0, date.today().isoformat())
        self.gantt_start_filter.grid(row=0, column=4, sticky="w", padx=6, pady=6)

        ttk.Label(controls, text="End Date").grid(row=0, column=5, sticky="e", padx=6, pady=6)
        if DateEntry is not None:
            self.gantt_end_filter = DateEntry(controls, width=14, date_pattern="yyyy-mm-dd")
            self.gantt_end_filter.set_date(date(2100, 12, 31))
        else:
            self.gantt_end_filter = ttk.Entry(controls, width=17)
            self.gantt_end_filter.insert(0, "2100-12-31")
        self.gantt_end_filter.grid(row=0, column=6, sticky="w", padx=6, pady=6)

        ttk.Button(controls, text="Refresh Chart", command=lambda: self._safe_execute(self.refresh_gantt)).grid(
            row=0,
            column=7,
            sticky="w",
            padx=6,
            pady=6,
        )
        controls.columnconfigure(8, weight=1)

        chart_frame = ttk.LabelFrame(self.tab_gantt, text="Project Gantt")
        chart_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)
        self.gantt_canvas = tk.Canvas(chart_frame, background="#ffffff", height=320, highlightthickness=1, highlightbackground="#cccccc")
        self.gantt_canvas.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        details_frame = ttk.LabelFrame(self.tab_gantt, text="Chart Details")
        details_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 8))
        columns = ("job_number", "address", "start_date", "end_date", "duration_days", "status")
        self.gantt_tree = ttk.Treeview(details_frame, columns=columns, show="headings", height=8)
        headings = ("Job Number", "Address", "Start Date", "End Date", "Duration (Days)", "Status")
        widths = (120, 320, 120, 120, 130, 130)
        for col, text, width in zip(columns, headings, widths):
            self.gantt_tree.heading(col, text=text)
            self.gantt_tree.column(col, width=width, anchor="w")

        scroll = ttk.Scrollbar(details_frame, orient="vertical", command=self.gantt_tree.yview)
        self.gantt_tree.configure(yscrollcommand=scroll.set)
        self.gantt_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        scroll.grid(row=0, column=1, sticky="ns", pady=6)

        self.tab_gantt.rowconfigure(1, weight=1)
        self.tab_gantt.rowconfigure(2, weight=1)
        self.tab_gantt.columnconfigure(0, weight=1)
        chart_frame.rowconfigure(0, weight=1)
        chart_frame.columnconfigure(0, weight=1)
        details_frame.rowconfigure(0, weight=1)
        details_frame.columnconfigure(0, weight=1)

    def build_map_tab(self) -> None:
        """Build map-like visualization and active location list using task data."""
        controls = ttk.LabelFrame(self.tab_map, text="Map Controls")
        controls.grid(row=0, column=0, sticky="ew", padx=10, pady=8)

        ttk.Label(controls, text="Map View built from active project locations.").grid(
            row=0,
            column=0,
            sticky="w",
            padx=6,
            pady=6,
        )
        ttk.Button(controls, text="Refresh Map", command=lambda: self._safe_execute(self.refresh_map)).grid(
            row=0,
            column=1,
            sticky="e",
            padx=6,
            pady=6,
        )
        controls.columnconfigure(0, weight=1)

        map_frame = ttk.LabelFrame(self.tab_map, text="Map View")
        map_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)
        self.map_canvas = tk.Canvas(
            map_frame,
            background="#f4f4f4",
            height=330,
            highlightthickness=1,
            highlightbackground="#cccccc",
        )
        self.map_canvas.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        active_frame = ttk.LabelFrame(self.tab_map, text="Active Location List")
        active_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 8))
        columns = ("marker", "job_number", "project_address", "status")
        self.map_tree = ttk.Treeview(active_frame, columns=columns, show="headings", height=9)
        headings = ("Marker", "Job Number", "Project Address", "Status")
        widths = (70, 120, 500, 140)
        for col, text, width in zip(columns, headings, widths):
            self.map_tree.heading(col, text=text)
            self.map_tree.column(col, width=width, anchor="w")

        scroll = ttk.Scrollbar(active_frame, orient="vertical", command=self.map_tree.yview)
        self.map_tree.configure(yscrollcommand=scroll.set)
        self.map_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        scroll.grid(row=0, column=1, sticky="ns", pady=6)

        self.tab_map.rowconfigure(1, weight=1)
        self.tab_map.rowconfigure(2, weight=1)
        self.tab_map.columnconfigure(0, weight=1)
        map_frame.rowconfigure(0, weight=1)
        map_frame.columnconfigure(0, weight=1)
        active_frame.rowconfigure(0, weight=1)
        active_frame.columnconfigure(0, weight=1)

    def refresh_all_views(self) -> None:
        """Refresh every dynamic control from current data."""
        self.refresh_project_tree()
        self.refresh_task_tree()
        self.refresh_report_filter_values()
        self.run_filtered_report()
        self.refresh_gantt()
        self.refresh_map()

    def _find_project_by_job(self, job_number: str) -> dict[str, object] | None:
        """Return project by job number or None when no match exists."""
        for project in self.data["projects"]:
            if str(project.get("job_number", "")).strip() == str(job_number).strip():
                return project
        return None

    def _find_project_by_id(self, project_id: str) -> dict[str, object] | None:
        """Return project by internal identifier or None if not found."""
        for project in self.data["projects"]:
            if str(project.get("id")) == project_id:
                return project
        return None

    def _find_task_by_id(self, task_id: str) -> dict[str, object] | None:
        """Return task by internal identifier or None if not found."""
        for task in self.data["tasks"]:
            if str(task.get("id")) == task_id:
                return task
        return None

    def _project_input(self) -> dict[str, object]:
        """Read, validate, and normalize project form values."""
        project_type = self.project_entries["Project Type"].get().strip()
        project_address = self.project_entries["Project Address"].get().strip()
        job_number = self.project_entries["Job Number"].get().strip()
        project_manager = self.project_entries["Project Manager"].get().strip()
        crew_name = self.project_entries["Crew Name"].get().strip()
        duration_days = safe_int(self.project_entries["Duration (Days)"].get(), "Duration (Days)")
        bid_price = safe_float(self.project_entries["Bid Price"].get(), "Bid Price")
        job_cost = safe_float(self.project_entries["Job Cost"].get(), "Job Cost")
        if duration_days < 0:
            raise ValueError("Duration (Days) cannot be negative.")
        if not all([project_type, project_address, job_number, project_manager, crew_name]):
            raise ValueError("All project fields except P & L are required.")
        return {
            "project_type": project_type,
            "project_address": project_address,
            "job_number": job_number,
            "project_manager": project_manager,
            "crew_name": crew_name,
            "duration_days": duration_days,
            "bid_price": bid_price,
            "job_cost": job_cost,
            "p_and_l": bid_price - job_cost,
        }

    def update_project_pl_preview(self) -> None:
        """Update read-only P & L preview value from Bid Price and Job Cost inputs."""
        try:
            bid_price = safe_float(self.project_entries["Bid Price"].get(), "Bid Price")
            job_cost = safe_float(self.project_entries["Job Cost"].get(), "Job Cost")
            value = format_currency(bid_price - job_cost)
        except ValueError:
            value = ""
        self.project_entries["P & L"].configure(state="normal")
        self.project_entries["P & L"].delete(0, tk.END)
        self.project_entries["P & L"].insert(0, value)
        self.project_entries["P & L"].configure(state="readonly")

    def add_project(self) -> None:
        """Add a new project from form data."""
        payload = self._project_input()
        if self._find_project_by_job(str(payload["job_number"])) is not None:
            raise ValueError("Job Number already exists.")
        payload["id"] = str(uuid.uuid4())
        self.data["projects"].append(payload)
        self.save_data()
        self.refresh_all_views()
        self.clear_project_form()
        messagebox.showinfo("Success", "Project added.")

    def update_project(self) -> None:
        """Update selected project with current form values."""
        if not self.selected_project_id:
            raise ValueError("Select a project to update.")
        project = self._find_project_by_id(self.selected_project_id)
        if project is None:
            raise ValueError("Selected project was not found.")
        payload = self._project_input()
        existing = self._find_project_by_job(str(payload["job_number"]))
        if existing is not None and str(existing.get("id")) != self.selected_project_id:
            raise ValueError("Job Number already exists for another project.")
        for key, value in payload.items():
            project[key] = value
        self.save_data()
        self.refresh_all_views()
        messagebox.showinfo("Success", "Project updated.")

    def remove_project(self) -> None:
        """Remove selected project and linked tasks."""
        if not self.selected_project_id:
            raise ValueError("Select a project to remove.")
        project = self._find_project_by_id(self.selected_project_id)
        if project is None:
            raise ValueError("Selected project was not found.")
        if not messagebox.askyesno("Confirm Remove", "Remove selected project and linked tasks?"):
            return
        job_number = str(project.get("job_number", ""))
        self.data["projects"] = [p for p in self.data["projects"] if str(p.get("id")) != self.selected_project_id]
        self.data["tasks"] = [t for t in self.data["tasks"] if str(t.get("job_number", "")).strip() != job_number]
        self.selected_project_id = None
        self.save_data()
        self.refresh_all_views()
        self.clear_project_form()
        messagebox.showinfo("Success", "Project removed.")

    def search_project(self) -> None:
        """Filter project tree by non-empty form fields."""
        filters = {
            "project_type": self.project_entries["Project Type"].get().strip().lower(),
            "project_address": self.project_entries["Project Address"].get().strip().lower(),
            "job_number": self.project_entries["Job Number"].get().strip().lower(),
            "project_manager": self.project_entries["Project Manager"].get().strip().lower(),
            "crew_name": self.project_entries["Crew Name"].get().strip().lower(),
        }

        filtered = []
        for project in self.data["projects"]:
            match = True
            for key, expected in filters.items():
                if expected and expected not in str(project.get(key, "")).lower():
                    match = False
                    break
            if match:
                filtered.append(project)
        self.refresh_project_tree(filtered)

    def clear_project_results(self) -> None:
        """Clear project form and reset tree to all projects."""
        self.clear_project_form()
        self.refresh_project_tree()

    def clear_project_form(self) -> None:
        """Clear project entry fields and active selection state."""
        self.selected_project_id = None
        for entry in self.project_entries.values():
            entry.configure(state="normal")
            entry.delete(0, tk.END)
        self.project_entries["Job Number"].insert(0, self._suggest_next_job_number())
        self.project_entries["P & L"].configure(state="readonly")

    def on_project_selected(self, _event) -> None:
        """Load selected project values into the form."""
        selected = self.project_tree.selection()
        if not selected:
            return
        project_id = selected[0]
        project = self._find_project_by_id(project_id)
        if project is None:
            return
        self.selected_project_id = project_id
        value_map = {
            "Project Type": project.get("project_type", ""),
            "Project Address": project.get("project_address", ""),
            "Job Number": project.get("job_number", ""),
            "Project Manager": project.get("project_manager", ""),
            "Crew Name": project.get("crew_name", ""),
            "Duration (Days)": project.get("duration_days", ""),
            "Bid Price": project.get("bid_price", 0.0),
            "Job Cost": project.get("job_cost", 0.0),
            "P & L": project.get("p_and_l", 0.0),
        }
        for label, entry in self.project_entries.items():
            entry.configure(state="normal")
            entry.delete(0, tk.END)
            value = value_map[label]
            if label in {"Bid Price", "Job Cost", "P & L"}:
                entry.insert(0, format_currency(float(value)).replace("$", ""))
            else:
                entry.insert(0, str(value))
            if label == "P & L":
                entry.configure(state="readonly")

    def refresh_project_tree(self, rows: list[dict[str, object]] | None = None) -> None:
        """Render provided project rows or all projects into the Project Management tree."""
        self.project_tree.delete(*self.project_tree.get_children())
        source = rows if rows is not None else self.data["projects"]
        for project in source:
            self.project_tree.insert(
                "",
                "end",
                iid=str(project.get("id")),
                values=(
                    project.get("project_type", ""),
                    project.get("project_address", ""),
                    project.get("job_number", ""),
                    project.get("project_manager", ""),
                    project.get("crew_name", ""),
                    project.get("duration_days", 0),
                    format_currency(float(project.get("bid_price", 0.0))),
                    format_currency(float(project.get("job_cost", 0.0))),
                    format_currency(float(project.get("p_and_l", 0.0))),
                ),
            )
        self._refresh_project_dropdown_values()
        self.refresh_report_filter_values()
        self.refresh_task_job_values()

    def _refresh_project_dropdown_values(self) -> None:
        """Update editable combobox option lists with starter and existing project values."""
        if not isinstance(self.project_entries.get("Project Type"), ttk.Combobox):
            return
        type_values = sorted(
            set(PROJECT_TYPE_OPTIONS).union(
                str(p.get("project_type", "")).strip() for p in self.data["projects"] if p.get("project_type")
            )
        )
        manager_values = sorted(
            set(PROJECT_MANAGER_OPTIONS).union(
                str(p.get("project_manager", "")).strip() for p in self.data["projects"] if p.get("project_manager")
            )
        )
        crew_values = sorted(
            set(CREW_OPTIONS).union(
                str(p.get("crew_name", "")).strip() for p in self.data["projects"] if p.get("crew_name")
            )
        )
        self.project_entries["Project Type"]["values"] = type_values
        self.project_entries["Project Manager"]["values"] = manager_values
        self.project_entries["Crew Name"]["values"] = crew_values

    def export_projects(self) -> None:
        """Export project table to CSV or XLSX."""
        file_path = filedialog.asksaveasfilename(
            title="Export Projects",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if not file_path:
            return
        headers = [
            "Project Type",
            "Project Address",
            "Job Number",
            "Project Manager",
            "Crew Name",
            "Duration (Days)",
            "Bid Price",
            "Job Cost",
            "P & L",
        ]
        rows = [
            [
                p.get("project_type", ""),
                p.get("project_address", ""),
                p.get("job_number", ""),
                p.get("project_manager", ""),
                p.get("crew_name", ""),
                p.get("duration_days", 0),
                p.get("bid_price", 0.0),
                p.get("job_cost", 0.0),
                p.get("p_and_l", 0.0),
            ]
            for p in self.data["projects"]
        ]
        self._write_table(file_path, headers, rows)
        messagebox.showinfo("Export Complete", "Project export complete.")

    def refresh_task_job_values(self) -> None:
        """Populate job number dropdown from existing projects."""
        job_numbers = sorted({str(p.get("job_number", "")).strip() for p in self.data["projects"] if p.get("job_number")})
        self.task_job_entry["values"] = job_numbers

    def _task_input(self) -> dict[str, object]:
        """Read, validate, and normalize Track Project form values."""
        task_name = self.task_entry.get().strip()
        description = self.description_entry.get().strip()
        job_number = self.task_job_entry.get().strip()
        start_date = self._get_date_value(self.start_date_entry)
        end_date = self._get_date_value(self.end_date_entry)
        status = self.status_combo.get().strip()
        if not all([task_name, job_number, start_date, end_date, status]):
            raise ValueError("Task, Job Number, Start Date, End Date, and Status are required.")
        if status not in TRACK_STATUSES:
            raise ValueError("Status must be one of the allowed options.")
        start_obj = parse_date(start_date, "Start Date")
        end_obj = parse_date(end_date, "End Date")
        if end_obj < start_obj:
            raise ValueError("End Date must be on or after Start Date.")
        return {
            "task": task_name,
            "description": description,
            "job_number": job_number,
            "start_date": start_date,
            "end_date": end_date,
            "status": status,
        }

    def assign_task(self) -> None:
        """Create a task assignment for a project."""
        payload = self._task_input()
        if self._find_project_by_job(str(payload["job_number"])) is None:
            raise ValueError("Job Number does not exist in Project Management.")
        payload["id"] = str(uuid.uuid4())
        self.data["tasks"].append(payload)
        self.save_data()
        self.refresh_all_views()
        self.clear_task_form()
        messagebox.showinfo("Success", "Task assigned.")

    def mark_selected_completed(self) -> None:
        """Update selected task status to Done."""
        if not self.selected_task_id:
            raise ValueError("Select a task row to mark completed.")
        task = self._find_task_by_id(self.selected_task_id)
        if task is None:
            raise ValueError("Selected task was not found.")
        task["status"] = "Done"
        self.save_data()
        self.refresh_all_views()
        messagebox.showinfo("Success", "Task marked as completed.")

    def clear_task_form(self) -> None:
        """Clear task form fields and reset default status."""
        self.selected_task_id = None
        self.task_entry.delete(0, tk.END)
        self.description_entry.delete(0, tk.END)
        self.task_job_entry.set("")
        self._set_date_value(self.start_date_entry, date.today().isoformat())
        self._set_date_value(self.end_date_entry, date.today().isoformat())
        self.status_combo.set("Upcoming")

    def on_task_selected(self, _event) -> None:
        """Populate task form when user selects a task child row."""
        selected = self.task_tree.selection()
        if not selected:
            return
        item_id = selected[0]
        task_id = self.task_row_lookup.get(item_id)
        if not task_id:
            return
        task = self._find_task_by_id(task_id)
        if task is None:
            return
        self.selected_task_id = task_id
        self.task_entry.delete(0, tk.END)
        self.task_entry.insert(0, str(task.get("task", "")))
        self.description_entry.delete(0, tk.END)
        self.description_entry.insert(0, str(task.get("description", "")))
        self.task_job_entry.set(str(task.get("job_number", "")))
        self._set_date_value(self.start_date_entry, str(task.get("start_date", date.today().isoformat())))
        self._set_date_value(self.end_date_entry, str(task.get("end_date", date.today().isoformat())))
        self.status_combo.set(str(task.get("status", "Upcoming")))

    def edit_selected_task(self) -> None:
        """Save edits from the task form to the selected task."""
        if not self.selected_task_id:
            raise ValueError("Select a task row to edit.")
        task = self._find_task_by_id(self.selected_task_id)
        if task is None:
            raise ValueError("Selected task was not found.")
        payload = self._task_input()
        if self._find_project_by_job(str(payload["job_number"])) is None:
            raise ValueError("Job Number does not exist in Project Management.")
        for key, value in payload.items():
            task[key] = value
        self.save_data()
        self.refresh_all_views()
        messagebox.showinfo("Success", "Task updated.")

    def refresh_task_tree(self) -> None:
        """Render grouped task tree by job number and project address."""
        self.task_tree.delete(*self.task_tree.get_children())
        self.task_row_lookup.clear()
        grouped: dict[tuple[str, str], list[dict[str, object]]] = {}
        for task in self.data["tasks"]:
            job_number = str(task.get("job_number", "")).strip()
            project = self._find_project_by_job(job_number)
            address = str(project.get("project_address", "Unknown Address")) if project else "Unknown Address"
            grouped.setdefault((job_number, address), []).append(task)

        for (job_number, address), tasks in sorted(grouped.items(), key=lambda item: item[0][0]):
            parent_iid = f"group::{job_number}::{address}"
            label = f"{job_number} | {address}"
            group_duration = sum(
                calculate_duration_days(str(task.get("start_date", "")), str(task.get("end_date", "")))
                for task in tasks
            )
            self.task_tree.insert("", "end", iid=parent_iid, text=label, values=("", "", "", "", "", group_duration, ""))
            for task in tasks:
                task_id = str(task.get("id"))
                row_iid = f"task::{task_id}"
                days = calculate_days_onsite(str(task.get("start_date", "")), str(task.get("end_date", "")))
                duration = calculate_duration_days(str(task.get("start_date", "")), str(task.get("end_date", "")))
                self.task_tree.insert(
                    parent_iid,
                    "end",
                    iid=row_iid,
                    text="",
                    values=(
                        task.get("task", ""),
                        task.get("job_number", ""),
                        task.get("start_date", ""),
                        task.get("end_date", ""),
                        days,
                        duration,
                        task.get("status", ""),
                    ),
                )
                self.task_row_lookup[row_iid] = task_id
            self.task_tree.item(parent_iid, open=True)

    def import_tasks(self) -> None:
        """Import tasks from CSV or XLSX."""
        file_path = filedialog.askopenfilename(
            title="Import Tasks",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if not file_path:
            return
        rows = self._read_table(file_path)
        required = {"Task", "Job Number", "Start Date", "End Date", "Status"}
        imported_count = 0
        skipped_count = 0
        for row in rows:
            if not required.issubset(row.keys()):
                raise ValueError("Import file is missing one or more required columns.")
            candidate = {
                "task": str(row.get("Task", "")).strip(),
                "description": str(row.get("Description", "")).strip(),
                "job_number": str(row.get("Job Number", "")).strip(),
                "start_date": str(row.get("Start Date", "")).strip(),
                "end_date": str(row.get("End Date", "")).strip(),
                "status": str(row.get("Status", "")).strip(),
            }
            try:
                self._validate_import_task(candidate)
                candidate["id"] = str(uuid.uuid4())
                self.data["tasks"].append(candidate)
                imported_count += 1
            except ValueError:
                skipped_count += 1
        self.save_data()
        self.refresh_all_views()
        messagebox.showinfo(
            "Import Complete",
            f"Imported: {imported_count}\nSkipped: {skipped_count}",
        )

    def _validate_import_task(self, candidate: dict[str, str]) -> None:
        """Validate a candidate task record from imported data."""
        required_values = (
            candidate.get("task", ""),
            candidate.get("job_number", ""),
            candidate.get("start_date", ""),
            candidate.get("end_date", ""),
            candidate.get("status", ""),
        )
        if not all(required_values):
            raise ValueError("Required imported task fields are missing.")
        if candidate["status"] not in TRACK_STATUSES:
            raise ValueError("Invalid status in import data.")
        if self._find_project_by_job(candidate["job_number"]) is None:
            raise ValueError("Imported task references unknown job number.")
        start_obj = parse_date(candidate["start_date"], "Start Date")
        end_obj = parse_date(candidate["end_date"], "End Date")
        if end_obj < start_obj:
            raise ValueError("End Date must be on or after Start Date.")

    def export_tasks(self) -> None:
        """Export tasks to CSV or XLSX."""
        file_path = filedialog.asksaveasfilename(
            title="Export Tasks",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if not file_path:
            return
        headers = ["Task", "Description", "Job Number", "Start Date", "End Date", "Status", "Days Onsite"]
        rows = []
        for task in self.data["tasks"]:
            rows.append(
                [
                    task.get("task", ""),
                    task.get("description", ""),
                    task.get("job_number", ""),
                    task.get("start_date", ""),
                    task.get("end_date", ""),
                    task.get("status", ""),
                    calculate_days_onsite(str(task.get("start_date", "")), str(task.get("end_date", ""))),
                ]
            )
        self._write_table(file_path, headers, rows)
        messagebox.showinfo("Export Complete", "Task export complete.")

    def save_task_template(self) -> None:
        """Save an empty task import template in CSV or XLSX format."""
        file_path = filedialog.asksaveasfilename(
            title="Save Task Template",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if not file_path:
            return
        headers = ["Task", "Description", "Job Number", "Start Date", "End Date", "Status"]
        self._write_table(file_path, headers, [])
        messagebox.showinfo("Template Saved", "Template file created.")

    def refresh_report_filter_values(self) -> None:
        """Update report tab dropdown values based on loaded project data."""
        jobs = sorted({str(p.get("job_number", "")).strip() for p in self.data["projects"] if p.get("job_number")})
        addresses = sorted({str(p.get("project_address", "")).strip() for p in self.data["projects"] if p.get("project_address")})
        self.report_job_combo["values"] = jobs
        self.report_address_combo["values"] = addresses

    def _collect_report_rows(self, active_only: bool, job_filter: str, address_filter: str) -> list[dict[str, object]]:
        """Build report rows as one row per project with task-derived schedule totals."""
        rows: list[dict[str, object]] = []
        for project in self.data["projects"]:
            job_number = str(project.get("job_number", ""))
            address = str(project.get("project_address", ""))
            if job_filter and job_filter.lower() not in job_number.lower():
                continue
            if address_filter and address_filter.lower() not in address.lower():
                continue

            project_tasks = [t for t in self.data["tasks"] if str(t.get("job_number", "")).strip() == job_number.strip()]
            if not project_tasks:
                row = {
                    "job_number": job_number,
                    "project_address": address,
                    "project_type": project.get("project_type", ""),
                    "start_date": "",
                    "end_date": "",
                    "days_onsite": 0,
                    "status": "Not Scheduled",
                    "bid_price": float(project.get("bid_price", 0.0)),
                    "job_cost": float(project.get("job_cost", 0.0)),
                    "p_and_l": float(project.get("p_and_l", 0.0)),
                }
                if not active_only:
                    rows.append(row)
                continue

            valid_tasks: list[dict[str, object]] = []
            for task in project_tasks:
                try:
                    parse_date(str(task.get("start_date", "")), "Start Date")
                    parse_date(str(task.get("end_date", "")), "End Date")
                    valid_tasks.append(task)
                except ValueError:
                    LOGGER.warning("Skipping report task with invalid dates id=%s", task.get("id"))
            if not valid_tasks:
                continue

            project_status = self._project_status_from_tasks(valid_tasks)
            if active_only and project_status == "Done":
                continue

            start_dates = [parse_date(str(task.get("start_date", "")), "Start Date") for task in valid_tasks]
            end_dates = [parse_date(str(task.get("end_date", "")), "End Date") for task in valid_tasks]
            days_total = sum(
                calculate_duration_days(str(task.get("start_date", "")), str(task.get("end_date", "")))
                for task in valid_tasks
            )
            rows.append(
                {
                    "job_number": job_number,
                    "project_address": address,
                    "project_type": project.get("project_type", ""),
                    "start_date": min(start_dates).isoformat(),
                    "end_date": max(end_dates).isoformat(),
                    "days_onsite": days_total,
                    "status": project_status,
                    "bid_price": float(project.get("bid_price", 0.0)),
                    "job_cost": float(project.get("job_cost", 0.0)),
                    "p_and_l": float(project.get("p_and_l", 0.0)),
                }
            )
        return rows

    def run_filtered_report(self) -> None:
        """Run report using filter controls and refresh summary."""
        job_filter = self.report_job_combo.get().strip()
        address_filter = self.report_address_combo.get().strip()
        rows = self._collect_report_rows(self.active_only_var.get(), job_filter, address_filter)
        self.report_rows = rows
        self._render_report_rows(rows)
        self._update_summary(rows)

    def show_all_active_projects(self) -> None:
        """Show all active project rows."""
        self.active_only_var.set(True)
        self.report_job_combo.set("")
        self.report_address_combo.set("")
        self.run_filtered_report()

    def show_project_history(self) -> None:
        """Show full project history including completed tasks."""
        self.active_only_var.set(False)
        self.report_job_combo.set("")
        self.report_address_combo.set("")
        self.run_filtered_report()

    def _render_report_rows(self, rows: list[dict[str, object]]) -> None:
        """Render report rows in the report tree."""
        self.report_tree.delete(*self.report_tree.get_children())
        for idx, row in enumerate(rows):
            self.report_tree.insert(
                "",
                "end",
                iid=f"report::{idx}",
                values=(
                    row["job_number"],
                    row["project_address"],
                    row["project_type"],
                    row["start_date"],
                    row["end_date"],
                    row["days_onsite"],
                    row["status"],
                    format_currency(float(row["bid_price"])),
                    format_currency(float(row["job_cost"])),
                    format_currency(float(row["p_and_l"])),
                ),
            )

    def _update_summary(self, rows: list[dict[str, object]]) -> None:
        """Update summary view section with totals and status counts."""
        total_projects = len(self.data["projects"])
        total_rows = len(rows)
        total_bid = sum(float(r["bid_price"]) for r in rows)
        total_cost = sum(float(r["job_cost"]) for r in rows)
        total_pl = sum(float(r["p_and_l"]) for r in rows)
        status_counts: dict[str, int] = {}
        for row in rows:
            status = str(row["status"])
            status_counts[status] = status_counts.get(status, 0) + 1

        lines = [
            f"Data Mode: {self.data_visibility}",
            f"Total Projects: {total_projects}",
            f"Report Rows: {total_rows}",
            f"Total Bid Price: {format_currency(total_bid)}",
            f"Total Job Cost: {format_currency(total_cost)}",
            f"Total P & L: {format_currency(total_pl)}",
            "",
            "Status Counts:",
        ]
        for status, count in sorted(status_counts.items()):
            lines.append(f"  - {status}: {count}")

        self.summary_text.configure(state="normal")
        self.summary_text.delete("1.0", tk.END)
        self.summary_text.insert(tk.END, "\n".join(lines))
        self.summary_text.configure(state="disabled")

    def export_reports(self) -> None:
        """Export current report rows to CSV or XLSX."""
        file_path = filedialog.asksaveasfilename(
            title="Export Report",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if not file_path:
            return
        headers = [
            "Job Number",
            "Project Address",
            "Project Type",
            "Start Date",
            "End Date",
            "Days Onsite",
            "Status",
            "Bid Price",
            "Job Cost",
            "P & L",
        ]
        rows = []
        for row in self.report_rows:
            rows.append(
                [
                    row["job_number"],
                    row["project_address"],
                    row["project_type"],
                    row["start_date"],
                    row["end_date"],
                    row["days_onsite"],
                    row["status"],
                    row["bid_price"],
                    row["job_cost"],
                    row["p_and_l"],
                ]
            )
        self._write_table(file_path, headers, rows)
        messagebox.showinfo("Export Complete", "Report export complete.")

    def refresh_gantt(self) -> None:
        """Refresh gantt chart and chart details from task schedule data."""
        self.gantt_tree.delete(*self.gantt_tree.get_children())
        self.gantt_canvas.delete("all")
        self._set_date_value(self.gantt_start_filter, date.today().isoformat())
        filter_start = parse_date(self._get_date_value(self.gantt_start_filter), "Gantt Start Date")
        filter_end = parse_date(self._get_date_value(self.gantt_end_filter), "Gantt End Date")
        if filter_end < filter_start:
            raise ValueError("Gantt End Date must be on or after Gantt Start Date.")
        status_filter = self.gantt_status_filter_var.get().strip()
        show_done = self.gantt_show_done_var.get()

        rows: list[dict[str, object]] = []
        for task in self.data["tasks"]:
            project = self._find_project_by_job(str(task.get("job_number", "")))
            if project is None:
                continue
            try:
                start = parse_date(str(task.get("start_date", "")), "Start Date")
                end = parse_date(str(task.get("end_date", "")), "End Date")
            except ValueError:
                LOGGER.warning("Skipping gantt row with invalid dates for task id=%s", task.get("id"))
                continue
            status_value = str(task.get("status", "")).strip()
            if not show_done and status_value == "Done":
                continue
            if status_filter and status_filter != "All Statuses" and status_value != status_filter:
                continue
            if end < filter_start or start > filter_end:
                continue
            duration = (end - start).days + 1
            rows.append(
                {
                    "job_number": task.get("job_number", ""),
                    "address": project.get("project_address", ""),
                    "start_date": task.get("start_date", ""),
                    "end_date": task.get("end_date", ""),
                    "duration_days": max(duration, 1),
                    "status": status_value,
                    "task_name": task.get("task", ""),
                }
            )
        self.gantt_rows = rows
        for idx, row in enumerate(rows):
            self.gantt_tree.insert(
                "",
                "end",
                iid=f"gantt::{idx}",
                values=(
                    row["job_number"],
                    row["address"],
                    row["start_date"],
                    row["end_date"],
                    row["duration_days"],
                    row["status"],
                ),
            )
        self._draw_gantt_canvas(rows)

    def _draw_gantt_canvas(self, rows: list[dict[str, object]]) -> None:
        """Draw a simplified gantt timeline on canvas based on task dates."""
        width = max(1020, int(self.gantt_canvas.winfo_width()))
        height = max(300, int(self.gantt_canvas.winfo_height()))
        self.gantt_canvas.configure(scrollregion=(0, 0, width, height))
        if not rows:
            self.gantt_canvas.create_text(
                width // 2,
                height // 2,
                text="No task schedule data available.",
                fill="#666666",
                font=("Segoe UI", 12),
            )
            return

        parsed = []
        for row in rows:
            start = parse_date(str(row["start_date"]), "Start Date")
            end = parse_date(str(row["end_date"]), "End Date")
            parsed.append((row, start, end))
        min_start = min(item[1] for item in parsed)
        max_end = max(item[2] for item in parsed)
        total_days = max((max_end - min_start).days, 1)

        left_pad = 240
        right_pad = 30
        top_pad = 30
        row_height = 28
        usable_width = max(width - left_pad - right_pad, 300)

        self.gantt_canvas.create_line(left_pad, top_pad - 10, left_pad, top_pad + row_height * len(parsed), fill="#999999")
        self.gantt_canvas.create_text(left_pad, 12, text=min_start.isoformat(), anchor="w", font=("Segoe UI", 9))
        self.gantt_canvas.create_text(width - right_pad, 12, text=max_end.isoformat(), anchor="e", font=("Segoe UI", 9))

        status_colors = {
            "Done": "#4CAF50",
            "In Progress": "#2196F3",
            "On Hold": "#FF9800",
            "Change Order": "#9C27B0",
            "Upcoming": "#546E7A",
            "Todo": "#78909C",
        }

        for idx, (row, start, end) in enumerate(parsed):
            y1 = top_pad + idx * row_height
            y2 = y1 + 18
            start_offset = (start - min_start).days
            end_offset = (end - min_start).days
            x1 = left_pad + int((start_offset / total_days) * usable_width)
            x2 = left_pad + int((end_offset / total_days) * usable_width)
            if x2 <= x1:
                x2 = x1 + 2
            color = status_colors.get(str(row["status"]), "#607D8B")
            label = f"{row['job_number']} | {row['task_name']}"
            self.gantt_canvas.create_text(10, y1 + 9, text=label, anchor="w", font=("Segoe UI", 9))
            self.gantt_canvas.create_rectangle(x1, y1, x2, y2, fill=color, outline="")
            self.gantt_canvas.create_text(x2 + 4, y1 + 9, text=str(row["status"]), anchor="w", font=("Segoe UI", 8))

    def refresh_map(self) -> None:
        """Refresh map canvas and active location list from active tasks."""
        self.map_tree.delete(*self.map_tree.get_children())
        self.map_canvas.delete("all")
        active_rows = []
        unique_jobs: dict[tuple[str, str], dict[str, object]] = {}
        for task in self.data["tasks"]:
            if str(task.get("status", "")) == "Done":
                continue
            project = self._find_project_by_job(str(task.get("job_number", "")))
            if project is None:
                continue
            job_number = str(task.get("job_number", "")).strip()
            address = str(project.get("project_address", "")).strip()
            if not address:
                continue
            key = (job_number, address)
            if key not in unique_jobs:
                unique_jobs[key] = {
                    "job_number": job_number,
                    "project_address": address,
                    "status": str(task.get("status", "")).strip(),
                }
        active_rows = list(unique_jobs.values())
        if not active_rows:
            self.map_canvas.create_text(
                500,
                160,
                text="No active locations available.",
                fill="#666666",
                font=("Segoe UI", 12),
            )
            return

        width = max(1020, int(self.map_canvas.winfo_width()))
        height = max(320, int(self.map_canvas.winfo_height()))
        self.map_canvas.create_rectangle(10, 10, width - 10, height - 10, outline="#b0bec5", fill="#e3f2fd")
        self.map_canvas.create_text(20, 20, anchor="nw", text="Map View (active project locations)", font=("Segoe UI", 11, "bold"))

        marker_letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        displayed_addresses: list[str] = []
        for idx, row in enumerate(active_rows):
            marker = marker_letters[idx % len(marker_letters)]
            displayed_addresses.append(f"{marker}. {row['project_address']}")
            self.map_tree.insert(
                "",
                "end",
                iid=f"map::{idx}",
                values=(marker, row["job_number"], row["project_address"], row["status"]),
            )

        columns = 2
        col_width = (width - 50) // columns
        base_y = 50
        for idx, text in enumerate(displayed_addresses):
            col = idx % columns
            row_index = idx // columns
            x = 25 + col * col_width
            y = base_y + row_index * 28
            self.map_canvas.create_oval(x, y, x + 16, y + 16, fill="#ef5350", outline="")
            self.map_canvas.create_text(x + 8, y + 8, text=text[0], fill="white", font=("Segoe UI", 8, "bold"))
            self.map_canvas.create_text(x + 22, y + 8, text=text[3:], anchor="w", font=("Segoe UI", 9))

    def _write_table(self, file_path: str, headers: list[str], rows: list[list[object]]) -> None:
        """Write tabular data to CSV or XLSX based on file extension."""
        _, ext = os.path.splitext(file_path.lower())
        if ext == ".csv":
            with open(file_path, "w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerow(headers)
                writer.writerows(rows)
            return
        if ext == ".xlsx":
            if Workbook is None:
                raise ValueError("openpyxl is required for .xlsx export.")
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Export"
            sheet.append(headers)
            for row in rows:
                sheet.append(row)
            workbook.save(file_path)
            return
        raise ValueError("Unsupported file extension.")

    def _read_table(self, file_path: str) -> list[dict[str, object]]:
        """Read CSV or XLSX tabular data and return rows as dictionaries."""
        _, ext = os.path.splitext(file_path.lower())
        if ext == ".csv":
            with open(file_path, "r", encoding="utf-8-sig", newline="") as file:
                reader = csv.DictReader(file)
                return [dict(row) for row in reader]
        if ext == ".xlsx":
            if load_workbook is None:
                raise ValueError("openpyxl is required for .xlsx import.")
            workbook = load_workbook(file_path)
            sheet = workbook.active
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                return []
            headers = [str(value).strip() if value is not None else "" for value in rows[0]]
            result: list[dict[str, object]] = []
            for raw_row in rows[1:]:
                mapped = {}
                for idx, header in enumerate(headers):
                    mapped[header] = raw_row[idx] if idx < len(raw_row) else None
                result.append(mapped)
            return result
        raise ValueError("Unsupported import file extension.")


def main() -> None:
    """Start the Project Management GUI."""
    root = tk.Tk()
    app = ProjectManagementApp(root)
    def _handle_tk_exception(exc_type, exc_value, exc_traceback) -> None:
        """Handle uncaught Tk callbacks with sanitized user notifications."""
        LOGGER.exception("Tk callback exception", exc_info=(exc_type, exc_value, exc_traceback))
        messagebox.showerror("Application Error", "An unexpected application error occurred.")

    root.report_callback_exception = _handle_tk_exception
    root.mainloop()


if __name__ == "__main__":
    main()
