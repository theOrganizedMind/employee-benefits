import csv
import json
import logging
import os
import tkinter as tk
from dataclasses import dataclass
from datetime import date, datetime
from tkinter import filedialog, messagebox, ttk
from uuid import uuid4

from employee import Employee

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import matplotlib.pyplot as plt
except ImportError:
    plt = None

try:
    from tkcalendar import DateEntry
except ImportError:
    DateEntry = None


LOGGER = logging.getLogger(__name__)
if not LOGGER.handlers:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRIVATE_CONFIG_PATH = os.path.join(BASE_DIR, "json_files", "employee_details.json")
PUBLIC_CONFIG_PATH = os.path.join(BASE_DIR, "employee_details.json")
DATA_STORE_PATH = os.path.join(BASE_DIR, "json_files", "employee_benefits_records.json")

DATE_FORMAT = "%Y-%m-%d"
DEFAULT_REMINDER_LOOKAHEAD_DAYS = 60
SPREADSHEET_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")
MIN_EMPLOYEE_AUTOSUGGEST_CHARS = 5
DEFAULT_VACATION_HOURS = 80.0
DEFAULT_SICK_HOURS = 40.0
DEFAULT_EMPLOYER_CONTRIBUTION = 300.0

TAB_KEYS = ("time_off", "performance", "deductions", "raises_bonuses", "admin")


def load_employee_details_seed():
    """Load base employee details from private path first, then public path."""
    config_path = PRIVATE_CONFIG_PATH if os.path.isfile(PRIVATE_CONFIG_PATH) else PUBLIC_CONFIG_PATH
    with open(config_path, encoding="utf-8") as config_file:
        return json.load(config_file)


def parse_date(value, field_name, required=False):
    """Parse date in YYYY-MM-DD format or return None for empty values."""
    cleaned = str(value).strip()
    if not cleaned:
        if required:
            raise ValueError(f"{field_name} is required.")
        return None
    try:
        return datetime.strptime(cleaned, DATE_FORMAT).date()
    except ValueError as error:
        raise ValueError(f"{field_name} must use YYYY-MM-DD format.") from error


def as_float(value, field_name):
    """Convert value to float with explicit field-specific error messages."""
    try:
        return float(str(value).strip())
    except ValueError as error:
        raise ValueError(f"{field_name} must be a number.") from error


def normalize_text(value):
    """Normalize free text from form fields before storage."""
    return str(value).replace("\x00", "").strip()


def sanitize_for_spreadsheet(value):
    """Prevent spreadsheet formula injection by escaping dangerous prefixes."""
    text = str(value)
    if text.startswith(SPREADSHEET_FORMULA_PREFIXES):
        return f"'{text}"
    return text


def calculate_duration_days(start_date, end_date):
    """Calculate inclusive duration in days between two dates."""
    if end_date < start_date:
        raise ValueError("End Date cannot be before Start Date.")
    return (end_date - start_date).days + 1


def annualized_pay(pay_type, pay):
    """Convert hourly or salary pay to annual compensation."""
    if pay_type == "Hourly":
        return round(
            (pay * Employee.REG_HOURS + pay * Employee.OVERTIME_MULTIPLIER * Employee.OT_HOURS) * Employee.WEEKS,
            2,
        )
    return round(pay, 2)


def record_year(date_value):
    """Extract year from YYYY-MM-DD date text."""
    parsed = parse_date(date_value, "Date")
    return parsed.year if parsed else None


def make_default_store():
    """Create default store using employee seed records as initial Admin rows."""
    seed = load_employee_details_seed()
    current_year = str(date.today().year)
    admin_records = []
    for seed_row in seed:
        latest_year = max((year for year in seed_row if year.isdigit()), default=None)
        latest = seed_row.get(latest_year, {}) if latest_year else {}
        pay = float(latest.get("salary", 0) or 0)
        admin_records.append(
            {
                "id": uuid4().hex,
                "employee_name": normalize_text(seed_row.get("name", "")),
                "hire_date": "",
                "termination_date": "",
                "pay_type": "Salary",
                "pay": pay,
                "available_vacation_hours": DEFAULT_VACATION_HOURS,
                "available_sick_hours": DEFAULT_SICK_HOURS,
                "notes": "",
                "active": True,
            }
        )
    return {
        "meta": {"last_annual_reset_year": current_year},
        "settings": {
            "default_vacation_hours": DEFAULT_VACATION_HOURS,
            "default_sick_hours": DEFAULT_SICK_HOURS,
            "default_employer_contribution": DEFAULT_EMPLOYER_CONTRIBUTION,
        },
        "yearly_totals": {},
        "admin": admin_records,
        "time_off": [],
        "performance": [],
        "deductions": [],
        "raises_bonuses": [],
    }


@dataclass
class TabWidgets:
    frame: ttk.Frame
    form_vars: dict
    form_widgets: dict
    field_types: dict
    tree: ttk.Treeview
    columns: list
    key: str


class EmployeeBenefitsApp:
    """Main tabbed UI for employee benefits management."""

    def __init__(self, root):
        """Initialize state, build UI, and populate all tab content."""
        self.root = root
        self.root.title("Employee Benefits Manager")
        self.root.geometry("1400x900")
        self.root.minsize(1120, 720)
        self.store = self._load_store()
        self.tabs = {}
        self.settings_widgets = {}
        self.reminder_widgets = {}
        self.chart_widgets = {}
        self._apply_annual_reset_if_needed()
        self._build_ui()
        self._refresh_all_tabs()

    def _load_store(self):
        """Load persisted records and normalize structure for current schema."""
        if not os.path.exists(DATA_STORE_PATH):
            store = make_default_store()
            self._save_store(store)
            return store

        with open(DATA_STORE_PATH, encoding="utf-8") as file:
            loaded = json.load(file)

        store = make_default_store()
        store["meta"] = loaded.get("meta", store["meta"])
        if not isinstance(store["meta"], dict):
            store["meta"] = {"last_annual_reset_year": str(date.today().year)}
        store["settings"] = loaded.get("settings", store["settings"])
        if not isinstance(store["settings"], dict):
            store["settings"] = {
                "default_vacation_hours": DEFAULT_VACATION_HOURS,
                "default_sick_hours": DEFAULT_SICK_HOURS,
                "default_employer_contribution": DEFAULT_EMPLOYER_CONTRIBUTION,
            }
        store["yearly_totals"] = loaded.get("yearly_totals", {})
        if not isinstance(store["yearly_totals"], dict):
            store["yearly_totals"] = {}

        for key in TAB_KEYS:
            loaded_rows = loaded.get(key, [])
            if isinstance(loaded_rows, list):
                store[key] = loaded_rows

        self._migrate_store(store)
        return store

    def _migrate_store(self, store):
        """Backfill legacy record fields to the current data model."""
        current_year = str(date.today().year)
        if "last_annual_reset_year" not in store["meta"]:
            store["meta"]["last_annual_reset_year"] = current_year

        store["settings"].setdefault("default_vacation_hours", DEFAULT_VACATION_HOURS)
        store["settings"].setdefault("default_sick_hours", DEFAULT_SICK_HOURS)
        store["settings"].setdefault("default_employer_contribution", DEFAULT_EMPLOYER_CONTRIBUTION)
        store["settings"]["default_vacation_hours"] = max(
            0.0, float(store["settings"]["default_vacation_hours"])
        )
        store["settings"]["default_sick_hours"] = max(
            0.0, float(store["settings"]["default_sick_hours"])
        )
        store["settings"]["default_employer_contribution"] = max(
            0.0, float(store["settings"]["default_employer_contribution"])
        )

        default_vacation_hours = self._get_setting_value("default_vacation_hours", store=store)
        default_sick_hours = self._get_setting_value("default_sick_hours", store=store)
        default_employer_contribution = self._get_setting_value("default_employer_contribution", store=store)

        for row in store["admin"]:
            row.setdefault("available_vacation_hours", default_vacation_hours)
            row.setdefault("available_sick_hours", default_sick_hours)
            row.setdefault("active", True)
            row.setdefault("notes", "")
            row.setdefault("hire_date", "")
            row.setdefault("termination_date", "")
            row.setdefault("pay_type", "Hourly")
            row.setdefault("pay", 0.0)

        for row in store["deductions"]:
            if "monthly_amount" not in row and "amount" in row:
                row["monthly_amount"] = row.get("amount", 0)
            row.setdefault("monthly_amount", 0.0)
            row.setdefault("employer_contribution", default_employer_contribution)
            row.setdefault("weekly_deduction", round((float(row["monthly_amount"]) - float(row["employer_contribution"])) / 4, 2))
            row.pop("amount", None)

    def _save_store(self, store=None):
        """Persist the active store JSON to disk."""
        active_store = store or self.store
        os.makedirs(os.path.dirname(DATA_STORE_PATH), exist_ok=True)
        with open(DATA_STORE_PATH, "w", encoding="utf-8") as file:
            json.dump(active_store, file, indent=2)

    def _apply_annual_reset_if_needed(self):
        """Reset annual sick/vacation balances once per new calendar year."""
        current_year = str(date.today().year)
        last_reset_year = str(self.store["meta"].get("last_annual_reset_year", current_year))
        if current_year == last_reset_year:
            return
        default_vacation_hours = self._get_setting_value("default_vacation_hours")
        default_sick_hours = self._get_setting_value("default_sick_hours")
        for row in self.store["admin"]:
            row["available_vacation_hours"] = default_vacation_hours
            row["available_sick_hours"] = default_sick_hours
        self.store["meta"]["last_annual_reset_year"] = current_year
        self._save_store()

    def _build_ui(self):
        """Create notebook tabs and all tab layouts."""
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self._build_time_off_tab(notebook)
        self._build_performance_tab(notebook)
        self._build_deductions_tab(notebook)
        self._build_raises_tab(notebook)
        self._build_admin_tab(notebook)
        self._build_reminders_tab(notebook)
        self._build_charts_tab(notebook)
        self._build_settings_tab(notebook)

    def _get_setting_value(self, key, store=None):
        """Return a non-negative float setting value for the provided key."""
        active_store = store or self.store
        default_map = {
            "default_vacation_hours": DEFAULT_VACATION_HOURS,
            "default_sick_hours": DEFAULT_SICK_HOURS,
            "default_employer_contribution": DEFAULT_EMPLOYER_CONTRIBUTION,
        }
        fallback = default_map[key]
        settings = active_store.get("settings", {})
        try:
            return max(0.0, float(settings.get(key, fallback)))
        except (TypeError, ValueError):
            return fallback

    def _add_tab(self, notebook, title, key, fields, columns, buttons):
        """Create one tab with form fields, buttons, and a treeview."""
        tab_frame = ttk.Frame(notebook)
        notebook.add(tab_frame, text=title)

        container = ttk.Frame(tab_frame, padding=10)
        container.pack(fill="both", expand=True)

        form_frame = ttk.Frame(container)
        form_frame.pack(fill="x", pady=(0, 8))
        form_vars = {}
        form_widgets = {}
        field_types = {}

        for index, field in enumerate(fields):
            label, field_key, field_type = field[:3]
            choices = field[3] if len(field) > 3 else None
            row = index // 2
            col_offset = (index % 2) * 3
            ttk.Label(form_frame, text=label).grid(row=row, column=col_offset, sticky="e", padx=(2, 5), pady=4)

            var = tk.StringVar()
            widget = self._build_form_widget(form_frame, var, field_type, choices, field_key, key)
            widget.grid(row=row, column=col_offset + 1, sticky="we", padx=(0, 14), pady=4)

            form_vars[field_key] = var
            form_widgets[field_key] = widget
            field_types[field_key] = field_type

            if field_type == "combo" and choices:
                var.set(choices[0])

        for column_index in range(6):
            form_frame.columnconfigure(column_index, weight=1 if column_index in (1, 4) else 0)

        button_row = ttk.Frame(container)
        button_row.pack(fill="x", pady=(0, 8))
        for button_text, callback in buttons:
            ttk.Button(button_row, text=button_text, command=callback).pack(side="left", padx=4)

        tree_frame = ttk.Frame(container)
        tree_frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=18)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=160, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        tab_widgets = TabWidgets(
            frame=tab_frame,
            form_vars=form_vars,
            form_widgets=form_widgets,
            field_types=field_types,
            tree=tree,
            columns=columns,
            key=key,
        )
        self.tabs[key] = tab_widgets
        tree.bind("<<TreeviewSelect>>", lambda _event, tab_key=key: self._populate_form_from_selection(tab_key))
        return tab_widgets

    def _build_form_widget(self, parent, var, field_type, choices, field_key, tab_key):
        """Create one form widget based on field type and tab context."""
        if field_type == "combo":
            return ttk.Combobox(parent, textvariable=var, values=choices or (), state="readonly", width=36)

        if field_type == "employee_combo":
            combo = ttk.Combobox(parent, textvariable=var, values=(), state="normal", width=36)
            include_inactive = tab_key == "admin"
            self._bind_employee_autosuggest(combo, var, include_inactive=include_inactive)
            return combo

        if field_type == "date":
            if DateEntry is None:
                return ttk.Entry(parent, textvariable=var, width=36)
            return DateEntry(
                parent,
                textvariable=var,
                date_pattern="yyyy-mm-dd",
                width=33,
            )

        return ttk.Entry(parent, textvariable=var, width=36)

    def _bind_employee_autosuggest(self, combo, variable, include_inactive):
        """Attach employee autosuggest behavior to a combobox."""
        def apply_suggestions(show_dropdown):
            """Update combobox values based on typed employee query."""
            query = normalize_text(variable.get())
            if len(query) < MIN_EMPLOYEE_AUTOSUGGEST_CHARS:
                combo.configure(values=())
                return
            suggestions = self._get_employee_suggestions(query, include_inactive=include_inactive)
            combo.configure(values=suggestions)
            if show_dropdown and suggestions:
                combo.event_generate("<Down>")

        combo.bind("<KeyRelease>", lambda _event: apply_suggestions(show_dropdown=True))
        combo.bind("<Button-1>", lambda _event: apply_suggestions(show_dropdown=False))
        combo.bind("<<ComboboxSelected>>", lambda _event: apply_suggestions(show_dropdown=False))

    def _get_employee_suggestions(self, query, include_inactive=False):
        """Return employee names matching query text with optional inactive rows."""
        query_folded = query.casefold()
        names = []
        seen = set()
        for row in self.store["admin"]:
            if not include_inactive and not row.get("active", True):
                continue
            name = normalize_text(row.get("employee_name", ""))
            if not name:
                continue
            key = name.casefold()
            if key in seen:
                continue
            if query_folded in key:
                seen.add(key)
                names.append(name)
        names.sort(key=str.casefold)
        return names

    def _build_time_off_tab(self, notebook):
        """Build the Time Off tab and its controls."""
        fields = [
            ("Employee Name", "employee_name", "employee_combo"),
            ("Type", "time_off_type", "combo", ["Vacation", "Sick", "Not Paid"]),
            ("Start Date", "start_date", "date"),
            ("End Date", "end_date", "date"),
            ("Notes", "notes", "entry"),
        ]
        columns = ["Year", "Employee Name", "Type", "Start Date", "End Date", "Duration", "Notes"]
        buttons = [
            ("Add", lambda: self._add_record("time_off")),
            ("Update", lambda: self._update_record("time_off")),
            ("Delete", lambda: self._delete_record("time_off")),
            ("Search", lambda: self._search_records("time_off")),
            ("Clear", lambda: self._clear_form("time_off", refresh=True)),
            ("Import", lambda: self._import_records("time_off")),
            ("Export", lambda: self._export_records("time_off")),
        ]
        self._add_tab(notebook, "Time Off", "time_off", fields, columns, buttons)

    def _build_performance_tab(self, notebook):
        """Build the Performance tab and its controls."""
        fields = [
            ("Employee Name", "employee_name", "employee_combo"),
            ("Review Date", "review_date", "date"),
            ("Performance Rating (1-5)", "rating", "entry"),
            ("Summary", "summary", "entry"),
            ("Notes", "notes", "entry"),
        ]
        columns = ["Employee Name", "Review Date", "Rating", "Summary", "Notes"]
        buttons = [
            ("Add", lambda: self._add_record("performance")),
            ("Update", lambda: self._update_record("performance")),
            ("Delete", lambda: self._delete_record("performance")),
            ("Search", lambda: self._search_records("performance")),
            ("Clear", lambda: self._clear_form("performance", refresh=True)),
            ("Import", lambda: self._import_records("performance")),
            ("Export", lambda: self._export_records("performance")),
        ]
        self._add_tab(notebook, "Performance", "performance", fields, columns, buttons)

    def _build_deductions_tab(self, notebook):
        """Build the Deductions tab and its controls."""
        fields = [
            ("Employee Name", "employee_name", "employee_combo"),
            ("Type", "deduction_type", "combo", ["BCBS", "Loan", "Child Support"]),
            ("Monthly Amount", "monthly_amount", "entry"),
            ("Employer Contribution", "employer_contribution", "entry"),
            ("Start Date", "start_date", "date"),
            ("Notes", "notes", "entry"),
        ]
        columns = [
            "Employee Name",
            "Type",
            "Monthly Amount",
            "Employer Contribution",
            "Weekly Deduction",
            "Start Date",
            "Notes",
        ]
        buttons = [
            ("Add", lambda: self._add_record("deductions")),
            ("Update", lambda: self._update_record("deductions")),
            ("Delete", lambda: self._delete_record("deductions")),
            ("Search", lambda: self._search_records("deductions")),
            ("Clear", lambda: self._clear_form("deductions", refresh=True)),
            ("Import", lambda: self._import_records("deductions")),
            ("Export", lambda: self._export_records("deductions")),
        ]
        self._add_tab(notebook, "Deductions", "deductions", fields, columns, buttons)

    def _build_raises_tab(self, notebook):
        """Build the Raises & Bonuses tab and its controls."""
        fields = [
            ("Employee Name", "employee_name", "employee_combo"),
            ("Type", "change_type", "combo", ["Raise", "Bonus"]),
            ("Amount", "amount", "entry"),
            ("Bonus Percent", "bonus_percent", "entry"),
            ("Effective Date", "effective_date", "date"),
            ("Notes", "notes", "entry"),
        ]
        columns = ["Employee Name", "Type", "Amount", "Bonus Percent", "Effective Date", "Calculated Value", "Notes"]
        buttons = [
            ("Add", lambda: self._add_record("raises_bonuses")),
            ("Update", lambda: self._update_record("raises_bonuses")),
            ("Delete", lambda: self._delete_record("raises_bonuses")),
            ("Search", lambda: self._search_records("raises_bonuses")),
            ("Clear", lambda: self._clear_form("raises_bonuses", refresh=True)),
            ("Import", lambda: self._import_records("raises_bonuses")),
            ("Export", lambda: self._export_records("raises_bonuses")),
        ]
        self._add_tab(notebook, "Raises & Bonuses", "raises_bonuses", fields, columns, buttons)

    def _build_admin_tab(self, notebook):
        """Build the Admin tab and its controls."""
        fields = [
            ("Employee Name", "employee_name", "employee_combo"),
            ("Hire Date", "hire_date", "date"),
            ("Pay Type", "pay_type", "combo", ["Hourly", "Salary"]),
            ("Termination Date", "termination_date", "date"),
            ("Pay", "pay", "entry"),
            ("Available Vacation (Hours)", "available_vacation_hours", "entry"),
            ("Notes", "notes", "entry"),
            ("Available Sick (Hours)", "available_sick_hours", "entry"),
        ]
        columns = [
            "Employee Name",
            "Hire Date",
            "Termination Date",
            "Pay Type",
            "Pay",
            "Available Vacation (Hours)",
            "Available Sick (Hours)",
            "Active",
            "Notes",
        ]
        buttons = [
            ("Add", lambda: self._add_record("admin")),
            ("Update", lambda: self._update_record("admin")),
            ("Delete", lambda: self._delete_record("admin")),
            ("Search", lambda: self._search_records("admin")),
            ("Clear", lambda: self._clear_form("admin", refresh=True)),
            ("Activate", lambda: self._set_employee_active(True)),
            ("Deactivate", lambda: self._set_employee_active(False)),
            ("Import", lambda: self._import_records("admin")),
            ("Export", lambda: self._export_records("admin")),
        ]
        self._add_tab(notebook, "Admin", "admin", fields, columns, buttons)

    def _build_settings_tab(self, notebook):
        """Build the Settings tab for editing default preset values."""
        tab_frame = ttk.Frame(notebook)
        notebook.add(tab_frame, text="Settings")
        container = ttk.Frame(tab_frame, padding=10)
        container.pack(fill="both", expand=True)

        form_frame = ttk.Frame(container)
        form_frame.pack(fill="x", pady=(0, 8))

        vacation_var = tk.StringVar()
        sick_var = tk.StringVar()
        employer_var = tk.StringVar()

        ttk.Label(form_frame, text="Default Vacation Hours").grid(row=0, column=0, sticky="e", padx=(2, 5), pady=4)
        ttk.Entry(form_frame, textvariable=vacation_var, width=28).grid(row=0, column=1, sticky="w", padx=(0, 14), pady=4)
        ttk.Label(form_frame, text="Default Sick Hours").grid(row=1, column=0, sticky="e", padx=(2, 5), pady=4)
        ttk.Entry(form_frame, textvariable=sick_var, width=28).grid(row=1, column=1, sticky="w", padx=(0, 14), pady=4)
        ttk.Label(form_frame, text="Default Employer Contribution").grid(row=2, column=0, sticky="e", padx=(2, 5), pady=4)
        ttk.Entry(form_frame, textvariable=employer_var, width=28).grid(row=2, column=1, sticky="w", padx=(0, 14), pady=4)

        form_frame.columnconfigure(1, weight=1)

        button_row = ttk.Frame(container)
        button_row.pack(fill="x", pady=(0, 8))
        ttk.Button(button_row, text="Save", command=self._save_settings).pack(side="left", padx=4)
        ttk.Button(button_row, text="Reset to Defaults", command=self._reset_settings_to_defaults).pack(side="left", padx=4)

        info_label = ttk.Label(
            container,
            text=(
                "These presets apply to new Admin/Deductions defaults and annual January 1 balance resets."
            ),
            wraplength=900,
            justify="left",
        )
        info_label.pack(anchor="w", pady=(4, 0))

        self.settings_widgets = {
            "default_vacation_hours": vacation_var,
            "default_sick_hours": sick_var,
            "default_employer_contribution": employer_var,
        }
        self._load_settings_into_form()

    def _load_settings_into_form(self):
        """Populate Settings tab fields from persisted preset values."""
        if not self.settings_widgets:
            return
        self.settings_widgets["default_vacation_hours"].set(
            f"{self._get_setting_value('default_vacation_hours'):.2f}"
        )
        self.settings_widgets["default_sick_hours"].set(
            f"{self._get_setting_value('default_sick_hours'):.2f}"
        )
        self.settings_widgets["default_employer_contribution"].set(
            f"{self._get_setting_value('default_employer_contribution'):.2f}"
        )

    def _save_settings(self):
        """Validate and persist updated preset values from the Settings tab."""
        try:
            default_vacation_hours = as_float(
                self.settings_widgets["default_vacation_hours"].get(),
                "Default Vacation Hours",
            )
            default_sick_hours = as_float(
                self.settings_widgets["default_sick_hours"].get(),
                "Default Sick Hours",
            )
            default_employer_contribution = as_float(
                self.settings_widgets["default_employer_contribution"].get(),
                "Default Employer Contribution",
            )
            if (
                default_vacation_hours < 0
                or default_sick_hours < 0
                or default_employer_contribution < 0
            ):
                raise ValueError("Settings values cannot be negative.")
        except ValueError as error:
            messagebox.showerror("Settings", str(error), parent=self.root)
            return

        self.store["settings"]["default_vacation_hours"] = round(default_vacation_hours, 2)
        self.store["settings"]["default_sick_hours"] = round(default_sick_hours, 2)
        self.store["settings"]["default_employer_contribution"] = round(default_employer_contribution, 2)
        try:
            self._save_store()
            self._load_settings_into_form()
            self._clear_form("deductions", refresh=False)
            self._clear_form("admin", refresh=False)
            self._refresh_tab("admin")
            messagebox.showinfo("Settings", "Preset values updated.", parent=self.root)
        except OSError as error:
            messagebox.showerror("Settings", str(error), parent=self.root)

    def _reset_settings_to_defaults(self):
        """Reset preset values in the form back to built-in defaults."""
        self.settings_widgets["default_vacation_hours"].set(f"{DEFAULT_VACATION_HOURS:.2f}")
        self.settings_widgets["default_sick_hours"].set(f"{DEFAULT_SICK_HOURS:.2f}")
        self.settings_widgets["default_employer_contribution"].set(f"{DEFAULT_EMPLOYER_CONTRIBUTION:.2f}")

    def _build_reminders_tab(self, notebook):
        """Build the Reminders tab and reminder results table."""
        tab_frame = ttk.Frame(notebook)
        notebook.add(tab_frame, text="Reminders")
        container = ttk.Frame(tab_frame, padding=10)
        container.pack(fill="both", expand=True)

        form_frame = ttk.Frame(container)
        form_frame.pack(fill="x", pady=(0, 8))
        days_var = tk.StringVar(value=str(DEFAULT_REMINDER_LOOKAHEAD_DAYS))
        ttk.Label(form_frame, text="Lookahead Days").grid(row=0, column=0, sticky="e", padx=(2, 5), pady=4)
        ttk.Entry(form_frame, textvariable=days_var, width=20).grid(row=0, column=1, sticky="w", padx=(0, 14), pady=4)

        button_row = ttk.Frame(container)
        button_row.pack(fill="x", pady=(0, 8))
        ttk.Button(button_row, text="Import", command=lambda: self._import_records("time_off")).pack(side="left", padx=4)
        ttk.Button(button_row, text="Export", command=self._export_reminders).pack(side="left", padx=4)
        ttk.Button(button_row, text="Refresh", command=self._refresh_reminders).pack(side="left", padx=4)

        tree_frame = ttk.Frame(container)
        tree_frame.pack(fill="both", expand=True)
        columns = ["Category", "Employee Name", "Start Date", "End Date", "Duration", "Type", "Notes"]
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=170, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self.reminder_widgets = {"tree": tree, "days_var": days_var}

    def _build_charts_tab(self, notebook):
        """Build the Charts tab, launch controls, and yearly summary table."""
        tab_frame = ttk.Frame(notebook)
        notebook.add(tab_frame, text="Charts")
        container = ttk.Frame(tab_frame, padding=10)
        container.pack(fill="both", expand=True)

        form_frame = ttk.Frame(container)
        form_frame.pack(fill="x", pady=(0, 8))
        chart_var = tk.StringVar(value="Payroll")
        ttk.Label(form_frame, text="Chart Type").grid(row=0, column=0, sticky="e", padx=(2, 5), pady=4)
        ttk.Combobox(
            form_frame,
            textvariable=chart_var,
            values=["Payroll", "Raises and Bonuses", "Deductions", "Time Off Gantt"],
            state="readonly",
            width=30,
        ).grid(row=0, column=1, sticky="w", padx=(0, 14), pady=4)

        button_row = ttk.Frame(container)
        button_row.pack(fill="x", pady=(0, 8))
        ttk.Button(button_row, text="Launch Chart", command=self._launch_chart).pack(side="left", padx=4)
        ttk.Button(button_row, text="Refresh", command=self._refresh_chart_summary).pack(side="left", padx=4)

        tree_frame = ttk.Frame(container)
        tree_frame.pack(fill="both", expand=True)
        columns = ["Year", "Category", "Metric", "Value"]
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=220, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self.chart_widgets = {"tree": tree, "chart_var": chart_var}

    def _refresh_all_tabs(self):
        """Refresh all treeviews and computed dashboard sections."""
        self._refresh_tab("time_off")
        self._refresh_tab("performance")
        self._refresh_tab("deductions")
        self._refresh_tab("raises_bonuses")
        self._refresh_tab("admin")
        self._refresh_reminders()
        self._refresh_chart_summary()

    def _refresh_tab(self, tab_key, records=None):
        """Refresh one tab treeview with optional filtered records."""
        tab = self.tabs[tab_key]
        tree = tab.tree
        tree.delete(*tree.get_children())
        rows = records if records is not None else self.store[tab_key]

        if tab_key == "time_off":
            rows = sorted(
                rows,
                key=lambda row: (
                    record_year(row.get("start_date", "")) or 0,
                    normalize_text(row.get("employee_name", "")).casefold(),
                    row.get("start_date", ""),
                ),
            )

        for row in rows:
            tree.insert("", "end", iid=row["id"], values=self._display_row(tab_key, row))

    def _display_row(self, tab_key, record):
        """Return treeview display tuple for one record by tab type."""
        if tab_key == "time_off":
            year = record_year(record["start_date"])
            return (
                str(year) if year else "",
                record["employee_name"],
                record["time_off_type"],
                record["start_date"],
                record["end_date"],
                str(record["duration_days"]),
                record["notes"],
            )
        if tab_key == "performance":
            return (
                record["employee_name"],
                record["review_date"],
                "⭐" * int(record["rating"]),
                record["summary"],
                record["notes"],
            )
        if tab_key == "deductions":
            return (
                record["employee_name"],
                record["deduction_type"],
                f"{record['monthly_amount']:.2f}",
                f"{record['employer_contribution']:.2f}",
                f"{record['weekly_deduction']:.2f}",
                record["start_date"],
                record["notes"],
            )
        if tab_key == "raises_bonuses":
            return (
                record["employee_name"],
                record["change_type"],
                f"{record['amount']:.2f}",
                f"{record['bonus_percent'] * 100:.2f}%",
                record["effective_date"],
                f"{record['calculated_value']:.2f}",
                record["notes"],
            )
        if tab_key == "admin":
            remaining = self._remaining_time_off_hours(record["employee_name"])
            return (
                record["employee_name"],
                record["hire_date"],
                record["termination_date"],
                record["pay_type"],
                f"{record['pay']:.2f}",
                f"{remaining['vacation']:.2f}",
                f"{remaining['sick']:.2f}",
                "Active" if record.get("active", True) else "Inactive",
                record["notes"],
            )
        raise ValueError(f"Unsupported tab key: {tab_key}")

    def _populate_form_from_selection(self, tab_key):
        """Populate form fields from the currently selected treeview row."""
        tab = self.tabs[tab_key]
        selected = tab.tree.selection()
        if not selected:
            return
        record = self._find_record(tab_key, selected[0])
        if not record:
            return

        for field_key in tab.form_vars:
            value = record.get(field_key, "")
            if field_key == "bonus_percent":
                value = f"{float(value) * 100:.2f}" if value else ""
            self._set_form_value(tab_key, field_key, value)

    def _get_form_values(self, tab_key):
        """Return normalized form values for the specified tab."""
        tab = self.tabs[tab_key]
        return {key: normalize_text(var.get()) for key, var in tab.form_vars.items()}

    def _set_form_value(self, tab_key, field_key, value):
        """Set one form field value for a tab."""
        tab = self.tabs[tab_key]
        tab.form_vars[field_key].set("" if value is None else str(value))

    def _find_record(self, tab_key, record_id):
        """Find a tab record by its unique ID."""
        for row in self.store[tab_key]:
            if row["id"] == record_id:
                return row
        return None

    def _add_record(self, tab_key):
        """Create and persist a new record from the tab form."""
        try:
            record = self._build_record_from_form(tab_key)
            record["id"] = uuid4().hex
            self.store[tab_key].append(record)
            self._save_store()
            self._clear_form(tab_key, refresh=False)
            self._refresh_all_tabs()
        except (ValueError, OSError) as error:
            LOGGER.warning("Add failed for %s: %s", tab_key, error)
            messagebox.showerror("Add failed", str(error), parent=self.root)

    def _update_record(self, tab_key):
        """Update the selected record using current form values."""
        tab = self.tabs[tab_key]
        selected = tab.tree.selection()
        if not selected:
            messagebox.showinfo("Update", "Select a row to update.", parent=self.root)
            return
        record_id = selected[0]
        try:
            updated = self._build_record_from_form(tab_key, ignore_id=record_id)
            existing = self._find_record(tab_key, record_id)
            if not existing:
                raise ValueError("The selected record no longer exists.")
            updated["id"] = existing["id"]
            index = self.store[tab_key].index(existing)
            self.store[tab_key][index] = updated
            self._save_store()
            self._clear_form(tab_key, refresh=False)
            self._refresh_all_tabs()
        except (ValueError, OSError) as error:
            LOGGER.warning("Update failed for %s: %s", tab_key, error)
            messagebox.showerror("Update failed", str(error), parent=self.root)

    def _delete_record(self, tab_key):
        """Delete the selected record after user confirmation."""
        tab = self.tabs[tab_key]
        selected = tab.tree.selection()
        if not selected:
            messagebox.showinfo("Delete", "Select a row to delete.", parent=self.root)
            return
        record = self._find_record(tab_key, selected[0])
        if not record:
            messagebox.showerror("Delete failed", "The selected record no longer exists.", parent=self.root)
            return
        if not messagebox.askyesno("Delete", "Delete selected record?", parent=self.root):
            return
        self.store[tab_key].remove(record)
        try:
            self._save_store()
            self._clear_form(tab_key, refresh=False)
            self._refresh_all_tabs()
        except OSError as error:
            LOGGER.warning("Delete save failed for %s: %s", tab_key, error)
            messagebox.showerror("Delete failed", str(error), parent=self.root)

    def _search_records(self, tab_key):
        """Filter tab records using non-empty form values as criteria."""
        criteria = {
            key: value.casefold()
            for key, value in self._get_form_values(tab_key).items()
            if value
        }
        if not criteria:
            self._refresh_tab(tab_key)
            return
        filtered = []
        for record in self.store[tab_key]:
            if all(criteria[key] in str(record.get(key, "")).casefold() for key in criteria):
                filtered.append(record)
        self._refresh_tab(tab_key, records=filtered)

    def _clear_form(self, tab_key, refresh):
        """Clear tab form values and reapply tab defaults."""
        tab = self.tabs[tab_key]
        for key in tab.form_vars:
            self._set_form_value(tab_key, key, "")

        defaults = {
            "time_off": {"time_off_type": "Vacation"},
            "deductions": {
                "deduction_type": "BCBS",
                "employer_contribution": f"{self._get_setting_value('default_employer_contribution'):.2f}",
            },
            "raises_bonuses": {"change_type": "Raise"},
            "admin": {
                "pay_type": "Hourly",
                "available_vacation_hours": f"{self._get_setting_value('default_vacation_hours'):.2f}",
                "available_sick_hours": f"{self._get_setting_value('default_sick_hours'):.2f}",
            },
        }
        for key, value in defaults.get(tab_key, {}).items():
            self._set_form_value(tab_key, key, value)

        tab.tree.selection_remove(tab.tree.selection())
        if refresh:
            self._refresh_tab(tab_key)

    def _set_employee_active(self, active):
        """Set active status for the selected admin employee row."""
        tab = self.tabs["admin"]
        selected = tab.tree.selection()
        if not selected:
            messagebox.showinfo("Admin", "Select an employee row first.", parent=self.root)
            return
        record = self._find_record("admin", selected[0])
        if not record:
            messagebox.showerror("Admin", "Selected employee no longer exists.", parent=self.root)
            return
        record["active"] = bool(active)
        try:
            self._save_store()
            self._refresh_all_tabs()
        except OSError as error:
            LOGGER.warning("Failed setting active status: %s", error)
            messagebox.showerror("Admin", str(error), parent=self.root)

    def _build_record_from_form(self, tab_key, ignore_id=None):
        """Validate and transform form input into one normalized record."""
        values = self._get_form_values(tab_key)

        if tab_key == "time_off":
            employee_name = values["employee_name"]
            if not employee_name:
                raise ValueError("Employee Name is required.")
            start_date = parse_date(values["start_date"], "Start Date", required=True)
            end_date = parse_date(values["end_date"], "End Date", required=True)
            duration_days = calculate_duration_days(start_date, end_date)
            record = {
                "employee_name": employee_name,
                "time_off_type": values["time_off_type"] or "Vacation",
                "start_date": start_date.strftime(DATE_FORMAT),
                "end_date": end_date.strftime(DATE_FORMAT),
                "duration_days": duration_days,
                "notes": values["notes"],
            }
            self._validate_time_off_availability(record, ignore_id=ignore_id)
            return record

        if tab_key == "performance":
            employee_name = values["employee_name"]
            if not employee_name:
                raise ValueError("Employee Name is required.")
            review_date = parse_date(values["review_date"], "Review Date", required=True)
            rating = int(as_float(values["rating"], "Performance Rating"))
            if rating < 1 or rating > 5:
                raise ValueError("Performance Rating must be between 1 and 5.")
            return {
                "employee_name": employee_name,
                "review_date": review_date.strftime(DATE_FORMAT),
                "rating": rating,
                "summary": values["summary"],
                "notes": values["notes"],
            }

        if tab_key == "deductions":
            employee_name = values["employee_name"]
            if not employee_name:
                raise ValueError("Employee Name is required.")
            monthly_amount = as_float(values["monthly_amount"], "Monthly Amount")
            employer_contribution = as_float(
                values["employer_contribution"] or str(self._get_setting_value("default_employer_contribution")),
                "Employer Contribution",
            )
            if monthly_amount < 0 or employer_contribution < 0:
                raise ValueError("Monthly Amount and Employer Contribution cannot be negative.")
            start_date = parse_date(values["start_date"], "Start Date", required=True)
            weekly_deduction = round((monthly_amount - employer_contribution) / 4, 2)
            return {
                "employee_name": employee_name,
                "deduction_type": values["deduction_type"] or "BCBS",
                "monthly_amount": round(monthly_amount, 2),
                "employer_contribution": round(employer_contribution, 2),
                "weekly_deduction": weekly_deduction,
                "start_date": start_date.strftime(DATE_FORMAT),
                "notes": values["notes"],
            }

        if tab_key == "raises_bonuses":
            employee_name = values["employee_name"]
            if not employee_name:
                raise ValueError("Employee Name is required.")
            effective_date = parse_date(values["effective_date"], "Effective Date", required=True)
            change_type = values["change_type"] or "Raise"
            amount = as_float(values["amount"] or "0", "Amount")
            bonus_percent = as_float(values["bonus_percent"] or "0", "Bonus Percent")
            if bonus_percent > 1:
                bonus_percent = bonus_percent / 100
            if amount < 0 or bonus_percent < 0:
                raise ValueError("Amount and Bonus Percent cannot be negative.")

            admin_record = self._get_admin_record(employee_name)
            if not admin_record:
                raise ValueError("Employee must be in Admin tab before adding raises or bonuses.")
            annual_pay = annualized_pay(admin_record["pay_type"], admin_record["pay"])

            if change_type == "Raise":
                if amount <= 0:
                    raise ValueError("Raise Amount must be greater than 0.")
                calculated_value = round(
                    (
                        amount * Employee.REG_HOURS
                        + amount * Employee.OVERTIME_MULTIPLIER * Employee.OT_HOURS
                    )
                    * Employee.WEEKS,
                    2,
                )
            else:
                if bonus_percent <= 0:
                    raise ValueError("Bonus Percent must be greater than 0 for Bonus type.")
                calculated_value = round(annual_pay * bonus_percent, 2)

            return {
                "employee_name": employee_name,
                "change_type": change_type,
                "amount": round(amount, 2),
                "bonus_percent": round(bonus_percent, 6),
                "effective_date": effective_date.strftime(DATE_FORMAT),
                "calculated_value": calculated_value,
                "notes": values["notes"],
            }

        if tab_key == "admin":
            employee_name = values["employee_name"]
            if not employee_name:
                raise ValueError("Employee Name is required.")
            hire_date = parse_date(values["hire_date"], "Hire Date")
            termination_date = parse_date(values["termination_date"], "Termination Date")
            if hire_date and termination_date and termination_date < hire_date:
                raise ValueError("Termination Date cannot be before Hire Date.")
            pay = as_float(values["pay"], "Pay")
            if pay < 0:
                raise ValueError("Pay cannot be negative.")
            available_vacation_hours = as_float(
                values["available_vacation_hours"] or str(self._get_setting_value("default_vacation_hours")),
                "Available Vacation (Hours)",
            )
            available_sick_hours = as_float(
                values["available_sick_hours"] or str(self._get_setting_value("default_sick_hours")),
                "Available Sick (Hours)",
            )
            if available_vacation_hours < 0 or available_sick_hours < 0:
                raise ValueError("Available hours cannot be negative.")
            return {
                "employee_name": employee_name,
                "hire_date": hire_date.strftime(DATE_FORMAT) if hire_date else "",
                "termination_date": termination_date.strftime(DATE_FORMAT) if termination_date else "",
                "pay_type": values["pay_type"] or "Hourly",
                "pay": round(pay, 2),
                "available_vacation_hours": round(available_vacation_hours, 2),
                "available_sick_hours": round(available_sick_hours, 2),
                "notes": values["notes"],
                "active": True,
            }

        raise ValueError(f"Unsupported tab key: {tab_key}")

    def _get_admin_record(self, employee_name):
        """Return admin record matching employee name, if present."""
        target = employee_name.casefold()
        for row in self.store["admin"]:
            if normalize_text(row.get("employee_name", "")).casefold() == target:
                return row
        return None

    def _remaining_time_off_hours(self, employee_name, exclude_record_id=None):
        """Calculate remaining vacation and sick hours for current year."""
        admin_record = self._get_admin_record(employee_name)
        if not admin_record:
            return {"vacation": 0.0, "sick": 0.0}

        current_year = date.today().year
        used_vacation = 0.0
        used_sick = 0.0
        for row in self.store["time_off"]:
            if exclude_record_id and row["id"] == exclude_record_id:
                continue
            if normalize_text(row.get("employee_name", "")).casefold() != employee_name.casefold():
                continue
            start = parse_date(row.get("start_date", ""), "Start Date")
            if not start or start.year != current_year:
                continue
            hours = float(row.get("duration_days", 0)) * 8
            if row.get("time_off_type") == "Vacation":
                used_vacation += hours
            if row.get("time_off_type") == "Sick":
                used_sick += hours
        return {
            "vacation": round(float(admin_record["available_vacation_hours"]) - used_vacation, 2),
            "sick": round(float(admin_record["available_sick_hours"]) - used_sick, 2),
        }

    def _validate_time_off_availability(self, time_off_record, ignore_id=None):
        """Validate active status and available balances for a time-off request."""
        employee_name = time_off_record["employee_name"]
        admin_record = self._get_admin_record(employee_name)
        if not admin_record:
            raise ValueError("Employee must be in Admin tab before adding time off.")
        if not admin_record.get("active", True):
            raise ValueError("Cannot add time off for an inactive employee.")
        remaining = self._remaining_time_off_hours(employee_name, exclude_record_id=ignore_id)
        requested_hours = float(time_off_record["duration_days"]) * 8
        if time_off_record["time_off_type"] == "Vacation" and remaining["vacation"] - requested_hours < 0:
            raise ValueError("Not enough Available Vacation Hours for this request.")
        if time_off_record["time_off_type"] == "Sick" and remaining["sick"] - requested_hours < 0:
            raise ValueError("Not enough Available Sick Hours for this request.")

    def _import_records(self, tab_key):
        """Import tab records from CSV/XLSX with per-row validation."""
        path = filedialog.askopenfilename(
            parent=self.root,
            title="Import file",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if not path:
            return
        try:
            rows = self._read_rows(path)
            imported = 0
            rejected = 0
            for row in rows:
                try:
                    self._apply_import_row(tab_key, row)
                    imported += 1
                except ValueError as row_error:
                    LOGGER.warning("Rejected import row for %s: %s", tab_key, row_error)
                    rejected += 1
            self._save_store()
            self._refresh_all_tabs()
            messagebox.showinfo("Import complete", f"Imported {imported} row(s). Rejected {rejected} row(s).", parent=self.root)
        except (OSError, ValueError, csv.Error) as error:
            LOGGER.warning("Import failed for %s: %s", tab_key, error)
            messagebox.showerror("Import failed", str(error), parent=self.root)

    def _read_rows(self, path):
        """Read CSV/XLSX data and return a list of dictionaries."""
        extension = os.path.splitext(path)[1].lower()
        if extension == ".csv":
            with open(path, newline="", encoding="utf-8-sig") as file:
                return list(csv.DictReader(file))
        if extension == ".xlsx":
            if openpyxl is None:
                raise ValueError("openpyxl is required for XLSX import. Install openpyxl first.")
            workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
            sheet = workbook.active
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                return []
            headers = [str(value).strip() if value is not None else "" for value in rows[0]]
            parsed = []
            for row in rows[1:]:
                parsed.append({headers[idx]: "" if cell is None else str(cell) for idx, cell in enumerate(row)})
            return parsed
        raise ValueError("Only .csv and .xlsx files are supported.")

    def _apply_import_row(self, tab_key, row):
        """Route one imported row to the tab-specific import builder."""
        if tab_key == "time_off":
            self.store["time_off"].append(self._build_import_time_off(row))
            return
        if tab_key == "performance":
            self.store["performance"].append(self._build_import_performance(row))
            return
        if tab_key == "deductions":
            self.store["deductions"].append(self._build_import_deductions(row))
            return
        if tab_key == "raises_bonuses":
            self.store["raises_bonuses"].append(self._build_import_raises(row))
            return
        if tab_key == "admin":
            self.store["admin"].append(self._build_import_admin(row))
            return
        raise ValueError(f"Import not supported for tab key: {tab_key}")

    def _build_import_time_off(self, row):
        """Build and validate one imported Time Off record."""
        employee_name = normalize_text(row.get("Employee Name", ""))
        start_date = parse_date(row.get("Start Date", ""), "Start Date", required=True)
        end_date = parse_date(row.get("End Date", ""), "End Date", required=True)
        record = {
            "id": uuid4().hex,
            "employee_name": employee_name,
            "time_off_type": normalize_text(row.get("Type", "Vacation")) or "Vacation",
            "start_date": start_date.strftime(DATE_FORMAT),
            "end_date": end_date.strftime(DATE_FORMAT),
            "duration_days": calculate_duration_days(start_date, end_date),
            "notes": normalize_text(row.get("Notes", "")),
        }
        if not employee_name:
            raise ValueError("Time Off import row missing Employee Name.")
        self._validate_time_off_availability(record)
        return record

    def _build_import_performance(self, row):
        """Build and validate one imported Performance record."""
        employee_name = normalize_text(row.get("Employee Name", ""))
        rating = int(as_float(row.get("Rating", "0"), "Rating"))
        review_date = parse_date(row.get("Review Date", ""), "Review Date", required=True)
        if not employee_name:
            raise ValueError("Performance import row missing Employee Name.")
        if rating < 1 or rating > 5:
            raise ValueError("Performance rating must be between 1 and 5.")
        return {
            "id": uuid4().hex,
            "employee_name": employee_name,
            "review_date": review_date.strftime(DATE_FORMAT),
            "rating": rating,
            "summary": normalize_text(row.get("Summary", "")),
            "notes": normalize_text(row.get("Notes", "")),
        }

    def _build_import_deductions(self, row):
        """Build and validate one imported Deductions record."""
        employee_name = normalize_text(row.get("Employee Name", ""))
        monthly_amount = as_float(row.get("Monthly Amount", row.get("Amount", "0")), "Monthly Amount")
        employer_contribution = as_float(
            row.get("Employer Contribution", str(self._get_setting_value("default_employer_contribution"))),
            "Employer Contribution",
        )
        start_date = parse_date(row.get("Start Date", ""), "Start Date", required=True)
        if not employee_name:
            raise ValueError("Deductions import row missing Employee Name.")
        return {
            "id": uuid4().hex,
            "employee_name": employee_name,
            "deduction_type": normalize_text(row.get("Type", "BCBS")) or "BCBS",
            "monthly_amount": round(monthly_amount, 2),
            "employer_contribution": round(employer_contribution, 2),
            "weekly_deduction": round((monthly_amount - employer_contribution) / 4, 2),
            "start_date": start_date.strftime(DATE_FORMAT),
            "notes": normalize_text(row.get("Notes", "")),
        }

    def _build_import_raises(self, row):
        """Build and validate one imported Raise/Bonus record."""
        employee_name = normalize_text(row.get("Employee Name", ""))
        change_type = normalize_text(row.get("Type", "Raise")) or "Raise"
        amount = as_float(row.get("Amount", "0"), "Amount")
        bonus_percent = as_float(row.get("Bonus Percent", "0"), "Bonus Percent")
        if bonus_percent > 1:
            bonus_percent = bonus_percent / 100
        effective_date = parse_date(row.get("Effective Date", ""), "Effective Date", required=True)
        if not employee_name:
            raise ValueError("Raises import row missing Employee Name.")
        admin_record = self._get_admin_record(employee_name)
        if not admin_record:
            raise ValueError("Employee must exist in Admin before importing Raise/Bonus.")
        annual_pay = annualized_pay(admin_record["pay_type"], admin_record["pay"])
        if change_type == "Raise":
            calculated_value = round(
                (
                    amount * Employee.REG_HOURS
                    + amount * Employee.OVERTIME_MULTIPLIER * Employee.OT_HOURS
                )
                * Employee.WEEKS,
                2,
            )
        else:
            calculated_value = round(annual_pay * bonus_percent, 2)
        return {
            "id": uuid4().hex,
            "employee_name": employee_name,
            "change_type": change_type,
            "amount": round(amount, 2),
            "bonus_percent": round(bonus_percent, 6),
            "effective_date": effective_date.strftime(DATE_FORMAT),
            "calculated_value": calculated_value,
            "notes": normalize_text(row.get("Notes", "")),
        }

    def _build_import_admin(self, row):
        """Build and validate one imported Admin record."""
        employee_name = normalize_text(row.get("Employee Name", ""))
        if not employee_name:
            raise ValueError("Employee Name is required.")
        hire_date = parse_date(row.get("Hire Date", ""), "Hire Date")
        termination_date = parse_date(row.get("Termination Date", ""), "Termination Date")
        if hire_date and termination_date and termination_date < hire_date:
            raise ValueError("Termination Date cannot be before Hire Date.")
        pay = as_float(row.get("Pay", "0"), "Pay")
        available_vacation_hours = as_float(
            row.get("Available Vacation (Hours)", str(self._get_setting_value("default_vacation_hours"))),
            "Available Vacation (Hours)",
        )
        available_sick_hours = as_float(
            row.get("Available Sick (Hours)", str(self._get_setting_value("default_sick_hours"))),
            "Available Sick (Hours)",
        )
        active_text = normalize_text(row.get("Active", "Active")).casefold()
        active = active_text not in {"inactive", "false", "0", "no"}
        return {
            "id": uuid4().hex,
            "employee_name": employee_name,
            "hire_date": hire_date.strftime(DATE_FORMAT) if hire_date else "",
            "termination_date": termination_date.strftime(DATE_FORMAT) if termination_date else "",
            "pay_type": normalize_text(row.get("Pay Type", "Hourly")) or "Hourly",
            "pay": round(pay, 2),
            "available_vacation_hours": round(available_vacation_hours, 2),
            "available_sick_hours": round(available_sick_hours, 2),
            "notes": normalize_text(row.get("Notes", "")),
            "active": active,
        }

    def _export_records(self, tab_key):
        """Export one tab's records to CSV/XLSX."""
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Export file",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if not path:
            return
        try:
            headers, rows = self._rows_for_export(tab_key)
            self._write_rows(path, headers, rows)
            messagebox.showinfo("Export complete", f"Exported {len(rows)} row(s).", parent=self.root)
        except (OSError, ValueError) as error:
            LOGGER.warning("Export failed for %s: %s", tab_key, error)
            messagebox.showerror("Export failed", str(error), parent=self.root)

    def _rows_for_export(self, tab_key):
        """Return export headers and rows for the specified tab."""
        if tab_key == "time_off":
            headers = ["Year", "Employee Name", "Type", "Start Date", "End Date", "Duration", "Notes"]
            rows = [
                {
                    "Year": record_year(row["start_date"]) or "",
                    "Employee Name": row["employee_name"],
                    "Type": row["time_off_type"],
                    "Start Date": row["start_date"],
                    "End Date": row["end_date"],
                    "Duration": row["duration_days"],
                    "Notes": row["notes"],
                }
                for row in self.store["time_off"]
            ]
            return headers, rows

        if tab_key == "performance":
            headers = ["Employee Name", "Review Date", "Rating", "Summary", "Notes"]
            rows = [
                {
                    "Employee Name": row["employee_name"],
                    "Review Date": row["review_date"],
                    "Rating": row["rating"],
                    "Summary": row["summary"],
                    "Notes": row["notes"],
                }
                for row in self.store["performance"]
            ]
            return headers, rows

        if tab_key == "deductions":
            headers = [
                "Employee Name",
                "Type",
                "Monthly Amount",
                "Employer Contribution",
                "Weekly Deduction",
                "Start Date",
                "Notes",
            ]
            rows = [
                {
                    "Employee Name": row["employee_name"],
                    "Type": row["deduction_type"],
                    "Monthly Amount": f"{row['monthly_amount']:.2f}",
                    "Employer Contribution": f"{row['employer_contribution']:.2f}",
                    "Weekly Deduction": f"{row['weekly_deduction']:.2f}",
                    "Start Date": row["start_date"],
                    "Notes": row["notes"],
                }
                for row in self.store["deductions"]
            ]
            return headers, rows

        if tab_key == "raises_bonuses":
            headers = ["Employee Name", "Type", "Amount", "Bonus Percent", "Effective Date", "Calculated Value", "Notes"]
            rows = [
                {
                    "Employee Name": row["employee_name"],
                    "Type": row["change_type"],
                    "Amount": f"{row['amount']:.2f}",
                    "Bonus Percent": f"{row['bonus_percent'] * 100:.2f}",
                    "Effective Date": row["effective_date"],
                    "Calculated Value": f"{row['calculated_value']:.2f}",
                    "Notes": row["notes"],
                }
                for row in self.store["raises_bonuses"]
            ]
            return headers, rows

        if tab_key == "admin":
            headers = [
                "Employee Name",
                "Hire Date",
                "Termination Date",
                "Pay Type",
                "Pay",
                "Available Vacation (Hours)",
                "Available Sick (Hours)",
                "Active",
                "Notes",
            ]
            rows = []
            for row in self.store["admin"]:
                remaining = self._remaining_time_off_hours(row["employee_name"])
                rows.append(
                    {
                        "Employee Name": row["employee_name"],
                        "Hire Date": row["hire_date"],
                        "Termination Date": row["termination_date"],
                        "Pay Type": row["pay_type"],
                        "Pay": f"{row['pay']:.2f}",
                        "Available Vacation (Hours)": f"{remaining['vacation']:.2f}",
                        "Available Sick (Hours)": f"{remaining['sick']:.2f}",
                        "Active": "Active" if row.get("active", True) else "Inactive",
                        "Notes": row["notes"],
                    }
                )
            return headers, rows

        raise ValueError(f"Export not supported for tab key: {tab_key}")

    def _write_rows(self, path, headers, rows):
        """Write export rows to CSV/XLSX with spreadsheet injection hardening."""
        extension = os.path.splitext(path)[1].lower()
        if extension == ".csv":
            with open(path, "w", newline="", encoding="utf-8") as file:
                writer = csv.DictWriter(file, fieldnames=headers)
                writer.writeheader()
                for row in rows:
                    safe_row = {header: sanitize_for_spreadsheet(row.get(header, "")) for header in headers}
                    writer.writerow(safe_row)
            return

        if extension == ".xlsx":
            if openpyxl is None:
                raise ValueError("openpyxl is required for XLSX export. Install openpyxl first.")
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(headers)
            for row in rows:
                sheet.append([sanitize_for_spreadsheet(row.get(header, "")) for header in headers])
            workbook.save(path)
            return

        raise ValueError("Only .csv and .xlsx exports are supported.")

    def _refresh_reminders(self):
        """Rebuild reminder rows for upcoming time-off and anniversaries."""
        if not self.reminder_widgets:
            return
        tree = self.reminder_widgets["tree"]
        tree.delete(*tree.get_children())

        try:
            lookahead_days = int(self.reminder_widgets["days_var"].get())
            if lookahead_days <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Reminders", "Lookahead days must be a positive integer.", parent=self.root)
            return

        today = date.today()
        limit = date.fromordinal(today.toordinal() + lookahead_days)
        rows = []

        for row in self.store["time_off"]:
            start = parse_date(row["start_date"], "Start Date")
            if start and today <= start <= limit:
                rows.append(
                    (
                        "Time Off",
                        row["employee_name"],
                        row["start_date"],
                        row["end_date"],
                        str(row["duration_days"]),
                        row["time_off_type"],
                        row["notes"],
                    )
                )

        for row in self.store["admin"]:
            hire_date = parse_date(row.get("hire_date", ""), "Hire Date")
            if not hire_date:
                continue
            anniversary = date(today.year, hire_date.month, hire_date.day)
            if anniversary < today:
                anniversary = date(today.year + 1, hire_date.month, hire_date.day)
            if anniversary <= limit:
                rows.append(
                    (
                        "Hire Anniversary",
                        row["employee_name"],
                        anniversary.strftime(DATE_FORMAT),
                        anniversary.strftime(DATE_FORMAT),
                        "1",
                        "Anniversary",
                        row.get("notes", ""),
                    )
                )

        rows.sort(key=lambda item: item[2])
        for values in rows:
            tree.insert("", "end", values=values)

    def _export_reminders(self):
        """Export reminder rows currently shown in the reminder table."""
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Export reminders",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if not path:
            return
        headers = ["Category", "Employee Name", "Start Date", "End Date", "Duration", "Type", "Notes"]
        rows = []
        for iid in self.reminder_widgets["tree"].get_children():
            values = self.reminder_widgets["tree"].item(iid, "values")
            rows.append(
                {
                    "Category": values[0],
                    "Employee Name": values[1],
                    "Start Date": values[2],
                    "End Date": values[3],
                    "Duration": values[4],
                    "Type": values[5],
                    "Notes": values[6],
                }
            )
        try:
            self._write_rows(path, headers, rows)
            messagebox.showinfo("Export complete", f"Exported {len(rows)} reminder row(s).", parent=self.root)
        except (OSError, ValueError) as error:
            LOGGER.warning("Reminder export failed: %s", error)
            messagebox.showerror("Export failed", str(error), parent=self.root)

    def _calculate_yearly_totals(self):
        """Compute payroll, raises/bonuses, and time-off totals by year."""
        totals = {}
        current_year = str(date.today().year)
        payroll_total = sum(
            annualized_pay(row["pay_type"], float(row["pay"]))
            for row in self.store["admin"]
            if row.get("active", True)
        )
        totals[current_year] = {
            "payroll": round(payroll_total, 2),
            "raises_bonuses": 0.0,
            "time_off_days": 0.0,
        }

        for row in self.store["raises_bonuses"]:
            year = str(record_year(row["effective_date"]))
            if year not in totals:
                totals[year] = {"payroll": 0.0, "raises_bonuses": 0.0, "time_off_days": 0.0}
            totals[year]["raises_bonuses"] += float(row["calculated_value"])

        for row in self.store["time_off"]:
            year = str(record_year(row["start_date"]))
            if year not in totals:
                totals[year] = {"payroll": 0.0, "raises_bonuses": 0.0, "time_off_days": 0.0}
            totals[year]["time_off_days"] += float(row["duration_days"])

        persisted = self.store.get("yearly_totals", {})
        for year, values in persisted.items():
            if year not in totals:
                totals[year] = {
                    "payroll": float(values.get("payroll", 0.0)),
                    "raises_bonuses": float(values.get("raises_bonuses", 0.0)),
                    "time_off_days": float(values.get("time_off_days", 0.0)),
                }
            else:
                if totals[year]["payroll"] == 0.0:
                    totals[year]["payroll"] = float(values.get("payroll", 0.0))
        return totals

    def _refresh_chart_summary(self):
        """Refresh yearly chart-summary tree and persist computed yearly totals."""
        if not self.chart_widgets:
            return
        yearly_totals = self._calculate_yearly_totals()
        self.store["yearly_totals"] = yearly_totals
        self._save_store()

        tree = self.chart_widgets["tree"]
        tree.delete(*tree.get_children())
        for year in sorted(yearly_totals.keys(), key=int):
            metrics = yearly_totals[year]
            tree.insert("", "end", values=(year, "Payroll", "Total Annual Active Payroll", f"{metrics['payroll']:,.2f}"))
            tree.insert("", "end", values=(year, "Raises & Bonuses", "Total Calculated Amount", f"{metrics['raises_bonuses']:,.2f}"))
            tree.insert("", "end", values=(year, "Time Off", "Total Scheduled Days", f"{metrics['time_off_days']:,.2f}"))

    def _launch_chart(self):
        """Launch the selected chart type in a separate matplotlib window."""
        if plt is None:
            messagebox.showerror("Charts", "matplotlib is required to launch charts.", parent=self.root)
            return
        chart_type = self.chart_widgets["chart_var"].get()
        if chart_type == "Payroll":
            self._launch_payroll_chart()
            return
        if chart_type == "Raises and Bonuses":
            self._launch_raise_bonus_chart()
            return
        if chart_type == "Deductions":
            self._launch_deductions_chart()
            return
        if chart_type == "Time Off Gantt":
            self._launch_time_off_gantt_chart()
            return
        messagebox.showerror("Charts", "Unsupported chart type selected.", parent=self.root)

    def _launch_payroll_chart(self):
        """Launch a yearly payroll bar chart."""
        yearly = self._calculate_yearly_totals()
        years = sorted(yearly.keys(), key=int)
        values = [yearly[year]["payroll"] for year in years]
        if not years:
            messagebox.showinfo("Charts", "No payroll data available.", parent=self.root)
            return
        plt.figure(figsize=(10, 6))
        plt.bar(years, values, color="#1f77b4")
        plt.title("Payroll by Year")
        plt.xlabel("Year")
        plt.ylabel("Total Annual Active Payroll")
        plt.tight_layout()
        plt.show()

    def _launch_raise_bonus_chart(self):
        """Launch a yearly raises-and-bonuses line chart."""
        yearly = self._calculate_yearly_totals()
        years = sorted(yearly.keys(), key=int)
        values = [yearly[year]["raises_bonuses"] for year in years]
        if not years:
            messagebox.showinfo("Charts", "No raise/bonus data available.", parent=self.root)
            return
        plt.figure(figsize=(10, 6))
        plt.plot(years, values, marker="o", color="#2ca02c")
        plt.title("Raises & Bonuses by Year")
        plt.xlabel("Year")
        plt.ylabel("Total Raises & Bonuses")
        plt.tight_layout()
        plt.show()

    def _launch_deductions_chart(self):
        """Launch a deductions composition pie chart."""
        totals = {}
        for row in self.store["deductions"]:
            key = row["deduction_type"]
            totals[key] = totals.get(key, 0.0) + float(row["weekly_deduction"])
        if not totals:
            messagebox.showinfo("Charts", "No deductions data available.", parent=self.root)
            return
        labels = list(totals.keys())
        values = [totals[label] for label in labels]
        plt.figure(figsize=(9, 6))
        plt.pie(values, labels=labels, autopct="%1.1f%%", startangle=90)
        plt.title("Weekly Deductions by Type")
        plt.tight_layout()
        plt.show()

    def _launch_time_off_gantt_chart(self):
        """Launch a Gantt-style chart for scheduled employee time off."""
        rows = sorted(
            self.store["time_off"],
            key=lambda item: (
                record_year(item["start_date"]) or 0,
                normalize_text(item["employee_name"]).casefold(),
                item["start_date"],
            ),
        )
        if not rows:
            messagebox.showinfo("Charts", "No time off data available.", parent=self.root)
            return
        figure, axis = plt.subplots(figsize=(12, 7))
        y_positions = []
        y_labels = []
        for index, row in enumerate(rows):
            start = parse_date(row["start_date"], "Start Date", required=True)
            end = parse_date(row["end_date"], "End Date", required=True)
            duration = calculate_duration_days(start, end)
            axis.barh(index, duration, left=start.toordinal(), color="#ff7f0e", edgecolor="black")
            axis.text(
                start.toordinal(),
                index,
                f" {row['employee_name']} ({row['time_off_type']})",
                va="center",
                fontsize=8,
            )
            y_positions.append(index)
            y_labels.append(f"{record_year(row['start_date'])} - {row['employee_name']}")
        axis.set_yticks(y_positions)
        axis.set_yticklabels(y_labels)
        axis.set_title("Employee Time Off Gantt Chart")
        axis.set_xlabel("Date")
        axis.xaxis_date()
        figure.autofmt_xdate()
        plt.tight_layout()
        plt.show()


def create_gui():
    """Create and run the tabbed employee benefits manager."""
    window = tk.Tk()
    app = EmployeeBenefitsApp(window)
    for tab_key in TAB_KEYS:
        app._clear_form(tab_key, refresh=False)
    window.mainloop()


if __name__ == "__main__":
    create_gui()
