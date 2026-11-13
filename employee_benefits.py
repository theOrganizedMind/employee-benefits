import csv
import importlib
import json
import logging
import os
import tkinter as tk
from dataclasses import dataclass
from datetime import date, datetime, timedelta
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
    mplcursors = importlib.import_module("mplcursors")
except ImportError:
    mplcursors = None

try:
    from tkcalendar import DateEntry
except ImportError:
    DateEntry = None


LOGGER = logging.getLogger(__name__)
if not LOGGER.handlers:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRIVATE_CONFIG_PATH = os.path.join(BASE_DIR, "json_files", "employee_details.json")
PUBLIC_CONFIG_PATH = os.path.join(BASE_DIR, "employee_details_public.json")
PRIVATE_DATA_STORE_PATH = os.path.join(BASE_DIR, "json_files", "employee_benefits_records.json")
PUBLIC_DATA_STORE_PATH = os.path.join(BASE_DIR, "employee_benefits_records_public.json")
DATA_STORE_PATH = PRIVATE_DATA_STORE_PATH

DATE_FORMAT = "%Y-%m-%d"
DEFAULT_REMINDER_LOOKAHEAD_DAYS = 60
SPREADSHEET_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")
MIN_EMPLOYEE_AUTOSUGGEST_CHARS = 5
DEFAULT_VACATION_HOURS = 80.0
DEFAULT_SICK_HOURS = 40.0
DEFAULT_EMPLOYER_CONTRIBUTION = 300.0

TAB_KEYS = ("time_off", "performance", "deductions", "raises_bonuses", "admin")
ANNUAL_WORK_HOURS = (Employee.REG_HOURS + (Employee.OVERTIME_MULTIPLIER * Employee.OT_HOURS)) * Employee.WEEKS


def load_employee_details_seed():
    """Load base employee details from private path first, then public path."""
    if os.path.isfile(PRIVATE_CONFIG_PATH):
        config_path = PRIVATE_CONFIG_PATH
    elif os.path.isfile(PUBLIC_CONFIG_PATH):
        config_path = PUBLIC_CONFIG_PATH
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


def nth_weekday_of_month(year, month, weekday, occurrence):
    """Return the nth weekday date for a month (weekday: Monday=0..Sunday=6)."""
    first_day = date(year, month, 1)
    offset = (weekday - first_day.weekday()) % 7
    day_number = 1 + offset + ((occurrence - 1) * 7)
    return date(year, month, day_number)


def last_weekday_of_month(year, month, weekday):
    """Return the last weekday date for a month (weekday: Monday=0..Sunday=6)."""
    if month == 12:
        next_month_first = date(year + 1, 1, 1)
    else:
        next_month_first = date(year, month + 1, 1)
    last_day = next_month_first - timedelta(days=1)
    offset = (last_day.weekday() - weekday) % 7
    return last_day - timedelta(days=offset)


def company_paid_holidays_for_year(year):
    """Return configured company paid holidays for one year."""
    thanksgiving = nth_weekday_of_month(year, 11, weekday=3, occurrence=4)
    holiday_dates = [
        ("New Year's Day", date(year, 1, 1)),
        ("Memorial Day", last_weekday_of_month(year, 5, weekday=0)),
        ("4th of July", date(year, 7, 4)),
        ("Labor Day", nth_weekday_of_month(year, 9, weekday=0, occurrence=1)),
        ("Thanksgiving Day", thanksgiving),
        ("Day After Thanksgiving", thanksgiving + timedelta(days=1)),
        ("Christmas Eve", date(year, 12, 24)),
        ("Christmas Day", date(year, 12, 25)),
    ]
    return holiday_dates


def annualized_pay(pay_type, pay):
    """Convert hourly or salary pay to annual compensation."""
    if pay_type == "Hourly":
        return round(
            (pay * Employee.REG_HOURS + pay * Employee.OVERTIME_MULTIPLIER * Employee.OT_HOURS) * Employee.WEEKS,
            2,
        )
    return round(pay, 2)


def hourly_rate_from_annual(annual_pay):
    """Convert annualized pay to hourly equivalent using configured work hours."""
    if ANNUAL_WORK_HOURS <= 0:
        return 0.0
    return round(float(annual_pay) / ANNUAL_WORK_HOURS, 2)


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
        self._raises_sync_in_progress = False
        self.data_store_path = self._resolve_data_store_path()
        self.store = self._load_store()
        self.tabs = {}
        self.settings_widgets = {}
        self.reminder_widgets = {}
        self.chart_widgets = {}
        self._hovertips_notice_shown = False
        self.gantt_widgets = {}
        self.gantt_rows = []
        self._apply_annual_reset_if_needed()
        self._build_ui()
        self._refresh_all_tabs()

    def _resolve_data_store_path(self):
        """Resolve private/public records path with explicit fallback support."""
        if DATA_STORE_PATH != PRIVATE_DATA_STORE_PATH:
            return DATA_STORE_PATH
        if os.path.exists(PRIVATE_DATA_STORE_PATH):
            return PRIVATE_DATA_STORE_PATH
        if os.path.exists(PUBLIC_DATA_STORE_PATH):
            return PUBLIC_DATA_STORE_PATH
        return PRIVATE_DATA_STORE_PATH

    def _load_store(self):
        """Load persisted records and normalize structure for current schema."""
        if not os.path.exists(self.data_store_path):
            store = make_default_store()
            self._save_store(store)
            return store

        with open(self.data_store_path, encoding="utf-8") as file:
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

        for row in store["time_off"]:
            row.setdefault("time_off_type", "Vacation")
            row.setdefault("notes", "")
            start_date = parse_date(row.get("start_date", ""), "Start Date")
            end_date = parse_date(row.get("end_date", ""), "End Date")
            if "duration_days" not in row:
                row["duration_days"] = (
                    calculate_duration_days(start_date, end_date)
                    if start_date and end_date
                    else 0
                )
            if "hours" not in row:
                row["hours"] = round(float(row.get("duration_days", 0) or 0.0) * 8, 2)

        for row in store["deductions"]:
            if "monthly_amount" not in row and "amount" in row:
                row["monthly_amount"] = row.get("amount", 0)
            row.setdefault("monthly_amount", 0.0)
            row.setdefault("employer_contribution", default_employer_contribution)
            row.setdefault("weekly_deduction", round((float(row["monthly_amount"]) - float(row["employer_contribution"])) / 4, 2))
            row.pop("amount", None)

        for row in store["raises_bonuses"]:
            self._migrate_raise_bonus_row(row, store["admin"])

        self._apply_pending_raises_to_admin(store=store)

    def _migrate_raise_bonus_row(self, row, admin_rows):
        """Backfill legacy raise/bonus records to the current schema."""
        change_type = normalize_text(row.get("change_type", "Raise")) or "Raise"
        employee_name = normalize_text(row.get("employee_name", ""))
        yearly_difference = float(row.get("yearly_difference", row.get("calculated_value", 0.0)) or 0.0)
        bonus_percent = float(row.get("bonus_percent", 0.0) or 0.0)
        if bonus_percent > 1:
            bonus_percent = bonus_percent / 100

        hourly_from = float(row.get("hourly_from_amount", 0.0) or 0.0)
        yearly_from = float(row.get("yearly_from_amount", 0.0) or 0.0)
        has_current_schema_values = (
            hourly_from > 0
            or yearly_from > 0
            or float(row.get("hourly_to_amount", 0.0) or 0.0) > 0
            or float(row.get("yearly_to_amount", 0.0) or 0.0) > 0
        )
        if employee_name and not has_current_schema_values:
            admin_record = self._find_admin_record(admin_rows, employee_name)
            if admin_record:
                hourly_from, yearly_from = self._compensation_snapshot(admin_record)
        if yearly_from == 0 and hourly_from > 0:
            yearly_from = annualized_pay("Hourly", hourly_from)
        if hourly_from == 0 and yearly_from > 0:
            hourly_from = hourly_rate_from_annual(yearly_from)

        bonus_amount = float(row.get("bonus_amount", 0.0) or 0.0)
        if change_type == "Bonus" and bonus_amount <= 0:
            bonus_amount = yearly_difference if yearly_difference > 0 else round(yearly_from * bonus_percent, 2)
        if change_type == "Raise":
            bonus_amount = 0.0
            bonus_percent = 0.0

        yearly_to = float(row.get("yearly_to_amount", 0.0) or 0.0)
        if yearly_to <= 0:
            if change_type == "Bonus":
                yearly_to = round(yearly_from + bonus_amount, 2)
            else:
                yearly_to = round(yearly_from + yearly_difference, 2)
        hourly_to = float(row.get("hourly_to_amount", 0.0) or 0.0)
        if hourly_to <= 0 and yearly_to > 0:
            hourly_to = hourly_rate_from_annual(yearly_to)

        yearly_difference = round(yearly_to - yearly_from, 2)
        payroll_increase_percent = 0.0 if yearly_from <= 0 else round((yearly_difference / yearly_from) * 100, 4)

        row["change_type"] = "Bonus" if change_type == "Bonus" else "Raise"
        row["hourly_from_amount"] = round(hourly_from, 2)
        row["hourly_to_amount"] = round(hourly_to, 2)
        row["yearly_from_amount"] = round(yearly_from, 2)
        row["yearly_to_amount"] = round(yearly_to, 2)
        row["yearly_difference"] = yearly_difference
        row["payroll_increase_percent"] = payroll_increase_percent
        row["bonus_amount"] = round(bonus_amount, 2)
        row["bonus_percent"] = round(bonus_percent, 6)
        row.setdefault("notes", "")
        row.setdefault("applied_to_admin", True)

        row.pop("amount", None)
        row.pop("calculated_value", None)

    def _apply_pending_raises_to_admin(self, store=None):
        """Apply effective-date-due raises to Admin pay one time per raise record."""
        active_store = store or self.store
        today = date.today()
        changed = False

        def parse_effective(row):
            """Parse effective date for sorting/application; invalid dates are ignored."""
            try:
                return parse_date(row.get("effective_date", ""), "Effective Date")
            except ValueError:
                return None

        rows = sorted(
            active_store["raises_bonuses"],
            key=lambda row: (
                parse_effective(row) or date.min,
                normalize_text(row.get("employee_name", "")).casefold(),
            ),
        )
        for row in rows:
            if normalize_text(row.get("change_type", "Raise")) != "Raise":
                row["applied_to_admin"] = True
                continue
            if row.get("applied_to_admin", False):
                continue
            effective_date = parse_effective(row)
            if not effective_date or effective_date > today:
                continue
            admin_record = self._find_admin_record(active_store["admin"], row.get("employee_name", ""))
            if not admin_record:
                continue
            pay_type = normalize_text(admin_record.get("pay_type", "Hourly")) or "Hourly"
            if pay_type == "Hourly":
                admin_record["pay"] = round(float(row.get("hourly_to_amount", admin_record.get("pay", 0.0)) or 0.0), 2)
            else:
                admin_record["pay"] = round(float(row.get("yearly_to_amount", admin_record.get("pay", 0.0)) or 0.0), 2)
            row["applied_to_admin"] = True
            changed = True
        return changed

    def _find_admin_record(self, admin_rows, employee_name):
        """Return one admin row from a supplied collection by employee name."""
        target = employee_name.casefold()
        for row in admin_rows:
            if normalize_text(row.get("employee_name", "")).casefold() == target:
                return row
        return None

    def _save_store(self, store=None):
        """Persist the active store JSON to disk."""
        active_store = store or self.store
        os.makedirs(os.path.dirname(self.data_store_path), exist_ok=True)
        with open(self.data_store_path, "w", encoding="utf-8") as file:
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
        self.notebook = notebook

        self._build_time_off_tab(notebook)
        self._build_performance_tab(notebook)
        self._build_deductions_tab(notebook)
        self._build_raises_tab(notebook)
        self._build_admin_tab(notebook)
        self._build_reminders_tab(notebook)
        self._build_charts_tab(notebook)
        self._build_gantt_tab(notebook)
        self._build_settings_tab(notebook)
        notebook.bind("<<NotebookTabChanged>>", self._on_notebook_tab_changed)

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
        if field_type == "readonly_entry":
            return ttk.Entry(parent, textvariable=var, state="readonly", width=36)

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

    def _wire_raises_form_behaviors(self):
        """Attach field synchronization for raises/bonuses form values."""
        tab = self.tabs["raises_bonuses"]
        tab.form_vars["employee_name"].trace_add("write", lambda *_: self._sync_raises_form_rates(source="employee"))
        tab.form_vars["change_type"].trace_add("write", lambda *_: self._sync_raises_form_rates(source="type"))
        tab.form_vars["hourly_to_amount"].trace_add("write", lambda *_: self._sync_raises_form_rates(source="hourly_to"))
        tab.form_vars["yearly_to_amount"].trace_add("write", lambda *_: self._sync_raises_form_rates(source="yearly_to"))

    def _sync_raises_form_rates(self, source):
        """Auto-populate raises from-amount fields from current admin compensation."""
        if self._raises_sync_in_progress:
            return
        form = self.tabs["raises_bonuses"].form_vars
        employee_name = normalize_text(form["employee_name"].get())
        admin_record = self._get_admin_record(employee_name) if employee_name else None
        if not admin_record:
            if source in {"employee", "type"}:
                for key in ("hourly_to_amount", "yearly_to_amount"):
                    form[key].set("")
            for key in ("hourly_from_amount", "yearly_from_amount"):
                form[key].set("")
            return

        pay_type = normalize_text(admin_record.get("pay_type", "Hourly")) or "Hourly"
        hourly_from, yearly_from = self._compensation_snapshot(admin_record)
        self._raises_sync_in_progress = True
        try:
            form["hourly_from_amount"].set(f"{hourly_from:.2f}")
            form["yearly_from_amount"].set(f"{yearly_from:.2f}")

            if source in {"employee", "type"}:
                form["hourly_to_amount"].set(f"{hourly_from:.2f}")
                form["yearly_to_amount"].set(f"{yearly_from:.2f}")

            if pay_type == "Hourly":
                hourly_text = normalize_text(form["hourly_to_amount"].get())
                yearly_text = normalize_text(form["yearly_to_amount"].get())
                if source in {"employee", "type", "hourly_to"} and hourly_text:
                    try:
                        hourly_to = as_float(hourly_text, "Hourly To Amount")
                    except ValueError:
                        hourly_to = None
                    if hourly_to is not None:
                        form["yearly_to_amount"].set(f"{annualized_pay('Hourly', hourly_to):.2f}")
                elif source == "yearly_to" and yearly_text:
                    try:
                        yearly_to = as_float(yearly_text, "Yearly To Amount")
                    except ValueError:
                        yearly_to = None
                    if yearly_to is not None:
                        form["hourly_to_amount"].set(f"{hourly_rate_from_annual(yearly_to):.2f}")
            else:
                if source in {"employee", "type"}:
                    form["hourly_to_amount"].set(f"{hourly_from:.2f}")

            if normalize_text(form["change_type"].get()) == "Raise":
                if not normalize_text(form["bonus_amount"].get()):
                    form["bonus_amount"].set("0.00")
                if not normalize_text(form["bonus_percent"].get()):
                    form["bonus_percent"].set("0.00")
        finally:
            self._raises_sync_in_progress = False

    def _compensation_snapshot(self, admin_record):
        """Return current hourly and yearly pay snapshot from one admin record."""
        pay_type = normalize_text(admin_record.get("pay_type", "Hourly")) or "Hourly"
        pay = float(admin_record.get("pay", 0.0) or 0.0)
        if pay_type == "Hourly":
            hourly_rate = self._normalize_hourly_rate(pay)
            yearly_amount = annualized_pay("Hourly", hourly_rate)
            return hourly_rate, yearly_amount
        yearly_amount = annualized_pay("Salary", pay)
        hourly_rate = hourly_rate_from_annual(yearly_amount)
        return hourly_rate, yearly_amount

    def _normalize_hourly_rate(self, pay_value):
        """Normalize stored hourly pay that may be persisted as annualized total."""
        pay = float(pay_value or 0.0)
        if pay <= 0:
            return 0.0
        # Legacy datasets may store annualized hourly totals in the hourly pay field.
        if pay > 1000:
            return hourly_rate_from_annual(pay)
        return round(pay, 2)

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
            ("Type", "time_off_type", "combo", ["", "Vacation", "Sick", "Not Paid", "Holiday"]),
            ("Start Date", "start_date", "date"),
            ("End Date", "end_date", "date"),
            ("Hours", "hours", "entry"),
            ("Notes", "notes", "entry"),
        ]
        columns = ["Year", "Employee Name", "Type", "Start Date", "End Date", "Duration", "Hours", "Notes"]
        buttons = [
            ("Add", lambda: self._add_record("time_off")),
            ("Update", lambda: self._update_record("time_off")),
            ("Delete", lambda: self._delete_record("time_off")),
            ("Search", lambda: self._search_records("time_off")),
            ("Clear", lambda: self._clear_form("time_off", refresh=True)),
            ("Import", lambda: self._import_records("time_off")),
            ("Export", lambda: self._export_records("time_off")),
            ("Template", lambda: self._export_template("time_off")),
        ]
        self._add_tab(notebook, "Time Off", "time_off", fields, columns, buttons)

    def _build_performance_tab(self, notebook):
        """Build the Performance tab and its controls."""
        fields = [
            ("Employee Name", "employee_name", "employee_combo"),
            ("Review Date", "review_date", "date"),
            ("Performance Rating (1-5)", "rating", "combo", ["1", "2", "3", "4", "5"]),
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
            ("Template", lambda: self._export_template("performance")),
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
            ("Template", lambda: self._export_template("deductions")),
        ]
        self._add_tab(notebook, "Deductions", "deductions", fields, columns, buttons)

    def _build_raises_tab(self, notebook):
        """Build the Raises & Bonuses tab and its controls."""
        fields = [
            ("Employee Name", "employee_name", "employee_combo"),
            ("Type", "change_type", "combo", ["Raise", "Bonus"]),
            ("Effective Date", "effective_date", "date"),
            ("Hourly From Amount", "hourly_from_amount", "readonly_entry"),
            ("Hourly To Amount", "hourly_to_amount", "entry"),
            ("Yearly From Amount", "yearly_from_amount", "readonly_entry"),
            ("Yearly To Amount", "yearly_to_amount", "entry"),
            ("Bonus Amount", "bonus_amount", "entry"),
            ("Bonus %", "bonus_percent", "entry"),
            ("Notes", "notes", "entry"),
        ]
        columns = [
            "Employee Name",
            "Type",
            "Effective Date",
            "Hourly From Amount",
            "Hourly To Amount",
            "Yearly From Amount",
            "Yearly To Amount",
            "Yearly Difference",
            "% Increase In Payroll",
            "Bonus Amount",
            "Bonus %",
        ]
        buttons = [
            ("Add", lambda: self._add_record("raises_bonuses")),
            ("Update", lambda: self._update_record("raises_bonuses")),
            ("Delete", lambda: self._delete_record("raises_bonuses")),
            ("Search", lambda: self._search_records("raises_bonuses")),
            ("Clear", lambda: self._clear_form("raises_bonuses", refresh=True)),
            ("Import", lambda: self._import_records("raises_bonuses")),
            ("Export", lambda: self._export_records("raises_bonuses")),
            ("Template", lambda: self._export_template("raises_bonuses")),
        ]
        self._add_tab(notebook, "Raises & Bonuses", "raises_bonuses", fields, columns, buttons)
        self._wire_raises_form_behaviors()

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
            ("Template", lambda: self._export_template("admin")),
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
            values=["Payroll", "Raises and Bonuses", "Deductions"],
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

    def _build_gantt_tab(self, notebook):
        """Build the Gantt Chart tab with timeline view and time-off detail tree."""
        tab_frame = ttk.Frame(notebook)
        notebook.add(tab_frame, text="Gantt Chart")
        container = ttk.Frame(tab_frame, padding=10)
        container.pack(fill="both", expand=True)

        control_frame = ttk.LabelFrame(container, text="Timeline Controls")
        control_frame.pack(fill="x", pady=(0, 8))

        status_label = ttk.Label(
            control_frame,
            text="Time off timeline from start date to end date.",
        )
        status_label.grid(row=0, column=0, sticky="w", padx=6, pady=6)

        ttk.Button(control_frame, text="Refresh Chart", command=self._refresh_gantt_tab).grid(
            row=0,
            column=1,
            sticky="e",
            padx=6,
            pady=6,
        )
        control_frame.columnconfigure(0, weight=1)

        chart_frame = ttk.LabelFrame(container, text="Gantt Chart View")
        chart_frame.pack(fill="both", expand=True, pady=(0, 8))

        gantt_canvas = tk.Canvas(
            chart_frame,
            background="#ffffff",
            highlightthickness=0,
        )
        gantt_y_scroll = ttk.Scrollbar(chart_frame, orient="vertical", command=gantt_canvas.yview)
        gantt_x_scroll = ttk.Scrollbar(chart_frame, orient="horizontal", command=gantt_canvas.xview)
        gantt_canvas.configure(yscrollcommand=gantt_y_scroll.set, xscrollcommand=gantt_x_scroll.set)
        gantt_canvas.grid(row=0, column=0, sticky="nsew")
        gantt_y_scroll.grid(row=0, column=1, sticky="ns")
        gantt_x_scroll.grid(row=1, column=0, sticky="ew")
        chart_frame.rowconfigure(0, weight=1)
        chart_frame.columnconfigure(0, weight=1)
        gantt_canvas.bind("<Configure>", self._on_gantt_canvas_resize)

        detail_frame = ttk.LabelFrame(container, text="Time Off")
        detail_frame.pack(fill="both", expand=True)
        columns = ("year", "employee_name", "time_off_type", "start_date", "end_date", "duration_days", "hours", "notes")
        detail_tree = ttk.Treeview(detail_frame, columns=columns, show="headings", height=8)
        detail_tree.heading("year", text="Year")
        detail_tree.heading("employee_name", text="Employee Name")
        detail_tree.heading("time_off_type", text="Type")
        detail_tree.heading("start_date", text="Start Date")
        detail_tree.heading("end_date", text="End Date")
        detail_tree.heading("duration_days", text="Duration")
        detail_tree.heading("hours", text="Hours")
        detail_tree.heading("notes", text="Notes")
        detail_tree.column("year", width=80, anchor="center")
        detail_tree.column("employee_name", width=190, anchor="w")
        detail_tree.column("time_off_type", width=95, anchor="center")
        detail_tree.column("start_date", width=105, anchor="center")
        detail_tree.column("end_date", width=105, anchor="center")
        detail_tree.column("duration_days", width=90, anchor="center")
        detail_tree.column("hours", width=80, anchor="center")
        detail_tree.column("notes", width=300, anchor="w")
        detail_scroll = ttk.Scrollbar(detail_frame, orient="vertical", command=detail_tree.yview)
        detail_tree.configure(yscrollcommand=detail_scroll.set)
        detail_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        detail_scroll.grid(row=0, column=1, sticky="ns", pady=6)
        detail_frame.rowconfigure(0, weight=1)
        detail_frame.columnconfigure(0, weight=1)

        self.gantt_widgets = {
            "tab_frame": tab_frame,
            "status_label": status_label,
            "canvas": gantt_canvas,
            "tree": detail_tree,
        }

    def _build_company_paid_holiday_rows(self, window_start, window_end):
        """Build synthetic company holiday time-off rows within a date range."""
        holiday_rows = []
        for year in range(window_start.year, window_end.year + 1):
            for holiday_name, holiday_date in company_paid_holidays_for_year(year):
                if holiday_date < window_start or holiday_date > window_end:
                    continue
                holiday_slug = holiday_name.lower().replace(" ", "-").replace("'", "")
                holiday_rows.append(
                    {
                        "id": f"holiday-{year}-{holiday_slug}",
                        "employee_name": holiday_name,
                        "time_off_type": "Holiday",
                        "start_date": holiday_date.strftime(DATE_FORMAT),
                        "end_date": holiday_date.strftime(DATE_FORMAT),
                        "duration_days": 1,
                        "hours": 8.0,
                        "notes": "Company Paid Holiday",
                    }
                )
        return holiday_rows

    def _build_time_off_gantt_rows(self):
        """Build normalized rows for records overlapping the next 30 days."""
        window_start = date.today()
        window_end = window_start + timedelta(days=29)
        rows = sorted(
            self.store["time_off"] + self._build_company_paid_holiday_rows(window_start, window_end),
            key=lambda item: (
                record_year(item["start_date"]) or 0,
                normalize_text(item["employee_name"]).casefold(),
                item["start_date"],
            ),
        )
        prepared = []
        for row in rows:
            start = parse_date(row["start_date"], "Start Date", required=True)
            end = parse_date(row["end_date"], "End Date", required=True)
            duration = calculate_duration_days(start, end)
            if end < window_start or start > window_end:
                continue
            chart_start = max(start, window_start)
            chart_end = min(end, window_end)
            prepared.append(
                {
                    "year": str(start.year),
                    "employee_name": row["employee_name"],
                    "time_off_type": row["time_off_type"],
                    "start_date": row["start_date"],
                    "end_date": row["end_date"],
                    "hours": self._time_off_hours(row),
                    "notes": row["notes"],
                    "start": start,
                    "end": end,
                    "chart_start": chart_start,
                    "chart_end": chart_end,
                    "chart_duration": calculate_duration_days(chart_start, chart_end),
                    "duration": duration,
                }
            )
        return prepared

    def _populate_gantt_time_off_tree(self):
        """Populate the Time Off detail tree under the Gantt chart."""
        tree = self.gantt_widgets["tree"]
        tree.delete(*tree.get_children())
        for row in self.gantt_rows:
            tree.insert(
                "",
                tk.END,
                values=(
                    row["year"],
                    row["employee_name"],
                    row["time_off_type"],
                    row["start_date"],
                    row["end_date"],
                    row["duration"],
                    f"{row['hours']:.2f}",
                    row["notes"],
                ),
            )

    def _draw_time_off_gantt_chart(self):
        """Render the Gantt canvas with one daily tick for the next 30 days."""
        canvas = self.gantt_widgets["canvas"]
        canvas.delete("all")
        window_start = date.today()
        window_end = window_start + timedelta(days=29)
        total_days = 30
        if not self.gantt_rows:
            canvas.create_text(
                200,
                80,
                text=f"No time off rows in the next 30 days ({window_start.isoformat()} to {window_end.isoformat()}).",
                fill="#444444",
                font=("Segoe UI", 11),
                anchor="w",
            )
            canvas.configure(scrollregion=(0, 0, 800, 220))
            return

        canvas_width = max(900, int(canvas.winfo_width()))
        left_margin = 265
        right_margin = 65
        top_margin = 56
        row_height = 28
        min_plot_width = total_days * 42
        plot_width = max(520, min_plot_width, canvas_width - left_margin - right_margin)
        plot_end_x = left_margin + plot_width
        content_height = top_margin + (len(self.gantt_rows) * row_height) + 40

        for day_offset in range(total_days):
            ratio = day_offset / max(1, total_days - 1)
            x = left_margin + int(plot_width * ratio)
            tick_date = window_start + timedelta(days=day_offset)
            canvas.create_line(x, top_margin - 20, x, content_height - 20, fill="#dddddd")
            canvas.create_text(
                x,
                top_margin - 28,
                text=tick_date.isoformat(),
                fill="#555555",
                font=("Segoe UI", 9),
                angle=45,
            )

        type_colors = {
            "Vacation": "#1d4ed8",
            "Sick": "#2a9d8f",
            "Not Paid": "#6b7280",
            "Holiday": "#ca8a04",
        }
        for index, row in enumerate(self.gantt_rows):
            y_top = top_margin + (index * row_height)
            y_bottom = y_top + 18
            y_mid = y_top + 9
            start_offset = (row["chart_start"] - window_start).days
            bar_days = row["chart_duration"]
            x0 = left_margin + int((start_offset / total_days) * plot_width)
            bar_width = max(2, int((bar_days / total_days) * plot_width))
            x1 = min(plot_end_x, x0 + bar_width)
            bar_color = type_colors.get(row["time_off_type"], "#ff7f0e")
            canvas.create_rectangle(x0, y_top, x1, y_bottom, fill=bar_color, outline="")
            canvas.create_text(
                10,
                y_mid,
                anchor="w",
                text=f"{row['employee_name']} ({row['time_off_type']})",
                font=("Segoe UI", 9),
                fill="#222222",
            )
            canvas.create_text(
                plot_end_x + 8,
                y_mid,
                anchor="w",
                text=f"{row['chart_duration']}d",
                font=("Segoe UI", 9),
                fill="#333333",
            )

        canvas.configure(scrollregion=(0, 0, plot_end_x + 95, content_height))

    def _on_gantt_canvas_resize(self, _event):
        """Redraw gantt chart when canvas dimensions change."""
        if self.gantt_rows:
            self._draw_time_off_gantt_chart()

    def _refresh_gantt_tab(self):
        """Reload Gantt tab rows and redraw canvas + detail tree."""
        if not self.gantt_widgets:
            return
        try:
            self.gantt_rows = self._build_time_off_gantt_rows()
        except ValueError as error:
            self.gantt_rows = []
            self.gantt_widgets["status_label"].config(text="Unable to load time off data for Gantt chart.")
            messagebox.showerror("Gantt Chart", str(error), parent=self.root)
            self._populate_gantt_time_off_tree()
            self._draw_time_off_gantt_chart()
            return

        self._populate_gantt_time_off_tree()
        self._draw_time_off_gantt_chart()
        if not self.gantt_rows:
            today = date.today()
            self.gantt_widgets["status_label"].config(
                text=f"No time off rows in the next 30 days ({today.isoformat()} to {(today + timedelta(days=29)).isoformat()})."
            )
            return
        today = date.today()
        self.gantt_widgets["status_label"].config(
            text=(
                f"Showing {len(self.gantt_rows)} time off row(s) from "
                f"{today.isoformat()} to {(today + timedelta(days=29)).isoformat()}."
            )
        )

    def _on_notebook_tab_changed(self, _event):
        """Refresh heavy visual tabs when selected."""
        if not self.gantt_widgets:
            return
        selected = self.notebook.select()
        if selected == str(self.gantt_widgets["tab_frame"]):
            self._refresh_gantt_tab()

    def _refresh_all_tabs(self):
        """Refresh all treeviews and computed dashboard sections."""
        self._refresh_tab("time_off")
        self._refresh_tab("performance")
        self._refresh_tab("deductions")
        self._refresh_tab("raises_bonuses")
        self._refresh_tab("admin")
        self._refresh_reminders()
        self._refresh_chart_summary()
        self._refresh_gantt_tab()

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
                f"{self._time_off_hours(record):.2f}",
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
                record["effective_date"],
                f"{record['hourly_from_amount']:.2f}",
                f"{record['hourly_to_amount']:.2f}",
                f"{record['yearly_from_amount']:.2f}",
                f"{record['yearly_to_amount']:.2f}",
                f"{record['yearly_difference']:.2f}",
                f"{record['payroll_increase_percent']:.2f}%",
                f"{record['bonus_amount']:.2f}",
                f"{record['bonus_percent'] * 100:.2f}%",
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
            if tab_key == "time_off":
                self._show_time_off_overlap_warning(record)
            record["id"] = uuid4().hex
            if tab_key == "raises_bonuses":
                record["applied_to_admin"] = False if record["change_type"] == "Raise" else True
            self.store[tab_key].append(record)
            if tab_key == "raises_bonuses":
                self._apply_pending_raises_to_admin()
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
            if tab_key == "time_off":
                self._show_time_off_overlap_warning(updated, ignore_id=record_id)
            existing = self._find_record(tab_key, record_id)
            if not existing:
                raise ValueError("The selected record no longer exists.")
            updated["id"] = existing["id"]
            if tab_key == "raises_bonuses":
                updated["applied_to_admin"] = False if updated["change_type"] == "Raise" else True
            index = self.store[tab_key].index(existing)
            self.store[tab_key][index] = updated
            if tab_key == "raises_bonuses":
                self._apply_pending_raises_to_admin()
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
            "time_off": {},
            "performance": {"rating": "3"},
            "deductions": {
                "deduction_type": "BCBS",
                "employer_contribution": f"{self._get_setting_value('default_employer_contribution'):.2f}",
            },
            "raises_bonuses": {"change_type": "Raise", "bonus_amount": "0.00", "bonus_percent": "0.00"},
            "admin": {
                "pay_type": "Hourly",
                "available_vacation_hours": f"{self._get_setting_value('default_vacation_hours'):.2f}",
                "available_sick_hours": f"{self._get_setting_value('default_sick_hours'):.2f}",
            },
        }
        for key, value in defaults.get(tab_key, {}).items():
            self._set_form_value(tab_key, key, value)

        tab.tree.selection_remove(tab.tree.selection())
        if tab_key == "raises_bonuses":
            self._sync_raises_form_rates(source="type")
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
            hours = as_float(values["hours"], "Hours")
            if hours <= 0:
                raise ValueError("Hours must be greater than zero.")
            record = {
                "employee_name": employee_name,
                "time_off_type": values["time_off_type"] or "Vacation",
                "start_date": start_date.strftime(DATE_FORMAT),
                "end_date": end_date.strftime(DATE_FORMAT),
                "duration_days": duration_days,
                "hours": round(hours, 2),
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
            return self._build_raise_bonus_record(
                employee_name=values["employee_name"],
                change_type=values["change_type"] or "Raise",
                effective_date_text=values["effective_date"],
                hourly_to_text=values["hourly_to_amount"],
                yearly_to_text=values["yearly_to_amount"],
                bonus_amount_text=values["bonus_amount"],
                bonus_percent_text=values["bonus_percent"],
                notes=values["notes"],
            )

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
        return self._find_admin_record(self.store["admin"], employee_name)

    def _build_raise_bonus_record(
        self,
        *,
        employee_name,
        change_type,
        effective_date_text,
        hourly_to_text,
        yearly_to_text,
        bonus_amount_text,
        bonus_percent_text,
        notes,
    ):
        """Build a normalized raise/bonus record from provided source values."""
        employee_name = normalize_text(employee_name)
        if not employee_name:
            raise ValueError("Employee Name is required.")
        admin_record = self._get_admin_record(employee_name)
        if not admin_record:
            raise ValueError("Employee must be in Admin tab before adding raises or bonuses.")

        effective_date = parse_date(effective_date_text, "Effective Date", required=True)
        change_type = "Bonus" if normalize_text(change_type) == "Bonus" else "Raise"
        pay_type = normalize_text(admin_record.get("pay_type", "Hourly")) or "Hourly"
        hourly_from, yearly_from = self._compensation_snapshot(admin_record)

        hourly_to_input = as_float(hourly_to_text or "0", "Hourly To Amount")
        yearly_to_input = as_float(yearly_to_text or "0", "Yearly To Amount")
        bonus_amount = as_float(bonus_amount_text or "0", "Bonus Amount")
        bonus_percent = as_float(bonus_percent_text or "0", "Bonus %")
        if bonus_percent > 1:
            bonus_percent = bonus_percent / 100

        if any(value < 0 for value in (hourly_to_input, yearly_to_input, bonus_amount, bonus_percent)):
            raise ValueError("Raise/bonus values cannot be negative.")

        if change_type == "Raise":
            if yearly_to_input <= 0 and hourly_to_input <= 0:
                raise ValueError("Raise requires Hourly To Amount or Yearly To Amount.")
            if pay_type == "Hourly":
                yearly_to_amount = yearly_to_input if yearly_to_input > 0 else annualized_pay("Hourly", hourly_to_input)
            else:
                yearly_to_amount = yearly_to_input if yearly_to_input > 0 else hourly_to_input
            if yearly_to_amount < yearly_from:
                raise ValueError("Raise Yearly To Amount must be greater than or equal to Yearly From Amount.")
            bonus_amount = 0.0
            bonus_percent = 0.0
        else:
            if bonus_amount <= 0 and bonus_percent <= 0:
                raise ValueError("Bonus requires Bonus Amount or Bonus %.")
            if bonus_amount <= 0:
                bonus_amount = round(yearly_from * bonus_percent, 2)
            if bonus_percent <= 0:
                bonus_percent = 0.0 if yearly_from <= 0 else bonus_amount / yearly_from
            yearly_to_amount = round(yearly_from + bonus_amount, 2)

        yearly_difference = round(yearly_to_amount - yearly_from, 2)
        hourly_to_amount = (
            round(hourly_to_input, 2)
            if pay_type == "Hourly" and hourly_to_input > 0
            else hourly_rate_from_annual(yearly_to_amount)
        )
        payroll_increase_percent = 0.0 if yearly_from <= 0 else round((yearly_difference / yearly_from) * 100, 4)

        return {
            "employee_name": employee_name,
            "change_type": change_type,
            "effective_date": effective_date.strftime(DATE_FORMAT),
            "hourly_from_amount": round(hourly_from, 2),
            "hourly_to_amount": round(hourly_to_amount, 2),
            "yearly_from_amount": round(yearly_from, 2),
            "yearly_to_amount": round(yearly_to_amount, 2),
            "yearly_difference": yearly_difference,
            "payroll_increase_percent": payroll_increase_percent,
            "bonus_amount": round(bonus_amount, 2),
            "bonus_percent": round(bonus_percent, 6),
            "notes": normalize_text(notes),
        }

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
            hours = self._time_off_hours(row)
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
        requested_hours = self._time_off_hours(time_off_record)
        if time_off_record["time_off_type"] == "Vacation" and remaining["vacation"] - requested_hours < 0:
            raise ValueError(
                f"{employee_name} does not currently have enough paid vacation hours available for this request."
            )
        if time_off_record["time_off_type"] == "Sick" and remaining["sick"] - requested_hours < 0:
            raise ValueError(
                f"{employee_name} does not currently have enough paid sick hours available for this request."
            )

    def _time_off_hours(self, time_off_record):
        """Return requested time-off hours with legacy fallback from duration days."""
        if "hours" in time_off_record and str(time_off_record["hours"]).strip():
            return round(as_float(time_off_record["hours"], "Hours"), 2)
        return round(float(time_off_record.get("duration_days", 0) or 0.0) * 8, 2)

    def _find_time_off_conflicts(self, time_off_record, ignore_id=None):
        """Find existing time-off rows that overlap with the supplied date range."""
        employee_name = normalize_text(time_off_record["employee_name"]).casefold()
        start = parse_date(time_off_record["start_date"], "Start Date", required=True)
        end = parse_date(time_off_record["end_date"], "End Date", required=True)
        overlaps = []
        for row in self.store["time_off"]:
            if ignore_id and row.get("id") == ignore_id:
                continue
            if normalize_text(row.get("employee_name", "")).casefold() != employee_name:
                continue
            existing_start = parse_date(row.get("start_date", ""), "Start Date", required=True)
            existing_end = parse_date(row.get("end_date", ""), "End Date", required=True)
            if start <= existing_end and end >= existing_start:
                overlaps.append(row)
        return overlaps

    def _show_time_off_overlap_warning(self, time_off_record, ignore_id=None):
        """Warn when an employee already has scheduled time off in the same period."""
        overlaps = self._find_time_off_conflicts(time_off_record, ignore_id=ignore_id)
        if not overlaps:
            return
        first = overlaps[0]
        messagebox.showwarning(
            "Time Off Overlap",
            (
                f"{time_off_record['employee_name']} already has time off from "
                f"{first['start_date']} to {first['end_date']}."
            ),
            parent=self.root,
        )

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
            if tab_key == "raises_bonuses":
                self._apply_pending_raises_to_admin()
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
        duration_days = calculate_duration_days(start_date, end_date)
        imported_hours = normalize_text(row.get("Hours", ""))
        hours = as_float(imported_hours, "Hours") if imported_hours else round(float(duration_days) * 8, 2)
        if hours <= 0:
            raise ValueError("Hours must be greater than zero.")
        record = {
            "id": uuid4().hex,
            "employee_name": employee_name,
            "time_off_type": normalize_text(row.get("Type", "Vacation")) or "Vacation",
            "start_date": start_date.strftime(DATE_FORMAT),
            "end_date": end_date.strftime(DATE_FORMAT),
            "duration_days": duration_days,
            "hours": round(hours, 2),
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
        record = self._build_raise_bonus_record(
            employee_name=row.get("Employee Name", ""),
            change_type=row.get("Type", "Raise"),
            effective_date_text=row.get("Effective Date", ""),
            hourly_to_text=row.get("Hourly To Amount", row.get("Hourly To", "0")),
            yearly_to_text=row.get("Yearly To Amount", row.get("Yearly To", "0")),
            bonus_amount_text=row.get("Bonus Amount", "0"),
            bonus_percent_text=row.get("Bonus %", row.get("Bonus Percent", "0")),
            notes=row.get("Notes", ""),
        )
        record["id"] = uuid4().hex
        record["applied_to_admin"] = False if record["change_type"] == "Raise" else True
        return record

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

    def _export_template(self, tab_key):
        """Export an empty CSV/XLSX template with tab import headers."""
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Export template",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if not path:
            return
        try:
            headers = self._template_headers(tab_key)
            self._write_rows(path, headers, [])
            messagebox.showinfo("Template complete", "Template exported.", parent=self.root)
        except (OSError, ValueError) as error:
            LOGGER.warning("Template export failed for %s: %s", tab_key, error)
            messagebox.showerror("Template export failed", str(error), parent=self.root)

    def _template_headers(self, tab_key):
        """Return template headers for one tab."""
        header_map = {
            "time_off": ["Employee Name", "Type", "Start Date", "End Date", "Hours", "Notes"],
            "performance": ["Employee Name", "Review Date", "Rating", "Summary", "Notes"],
            "deductions": [
                "Employee Name",
                "Type",
                "Monthly Amount",
                "Employer Contribution",
                "Start Date",
                "Notes",
            ],
            "raises_bonuses": [
                "Employee Name",
                "Type",
                "Effective Date",
                "Hourly To Amount",
                "Yearly To Amount",
                "Bonus Amount",
                "Bonus %",
                "Notes",
            ],
            "admin": [
                "Employee Name",
                "Hire Date",
                "Termination Date",
                "Pay Type",
                "Pay",
                "Available Vacation (Hours)",
                "Available Sick (Hours)",
                "Active",
                "Notes",
            ],
        }
        if tab_key not in header_map:
            raise ValueError(f"Template not supported for tab key: {tab_key}")
        return header_map[tab_key]

    def _rows_for_export(self, tab_key):
        """Return export headers and rows for the specified tab."""
        if tab_key == "time_off":
            headers = ["Year", "Employee Name", "Type", "Start Date", "End Date", "Duration", "Hours", "Notes"]
            rows = [
                {
                    "Year": record_year(row["start_date"]) or "",
                    "Employee Name": row["employee_name"],
                    "Type": row["time_off_type"],
                    "Start Date": row["start_date"],
                    "End Date": row["end_date"],
                    "Duration": row["duration_days"],
                    "Hours": f"{self._time_off_hours(row):.2f}",
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
            headers = [
                "Employee Name",
                "Type",
                "Effective Date",
                "Hourly From Amount",
                "Hourly To Amount",
                "Yearly From Amount",
                "Yearly To Amount",
                "Yearly Difference",
                "% Increase In Payroll",
                "Bonus Amount",
                "Bonus %",
                "Notes",
            ]
            rows = [
                {
                    "Employee Name": row["employee_name"],
                    "Type": row["change_type"],
                    "Effective Date": row["effective_date"],
                    "Hourly From Amount": f"{row['hourly_from_amount']:.2f}",
                    "Hourly To Amount": f"{row['hourly_to_amount']:.2f}",
                    "Yearly From Amount": f"{row['yearly_from_amount']:.2f}",
                    "Yearly To Amount": f"{row['yearly_to_amount']:.2f}",
                    "Yearly Difference": f"{row['yearly_difference']:.2f}",
                    "% Increase In Payroll": f"{row['payroll_increase_percent']:.2f}",
                    "Bonus Amount": f"{row['bonus_amount']:.2f}",
                    "Bonus %": f"{row['bonus_percent'] * 100:.2f}",
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
        """Rebuild reminder rows for upcoming time-off, holidays, and anniversaries."""
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

        for row in self.store["time_off"] + self._build_company_paid_holiday_rows(today, limit):
            start = parse_date(row["start_date"], "Start Date")
            if start and today <= start <= limit:
                category = "Company Holiday" if row["time_off_type"] == "Holiday" else "Time Off"
                rows.append(
                    (
                        category,
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
        """Compute payroll, raises/bonuses, time-off, and deductions totals by year."""
        totals = {}

        def ensure_year(year):
            totals.setdefault(
                year,
                {
                    "payroll": 0.0,
                    "raises_bonuses": 0.0,
                    "time_off_days": 0.0,
                    "deductions_monthly": 0.0,
                    "deductions_weekly": 0.0,
                },
            )

        current_year = str(date.today().year)
        ensure_year(current_year)
        totals[current_year]["payroll"] = round(
            sum(
                annualized_pay(row["pay_type"], float(row["pay"]))
                for row in self.store["admin"]
                if row.get("active", True)
            ),
            2,
        )

        for row in self.store["raises_bonuses"]:
            year = str(record_year(row["effective_date"]))
            ensure_year(year)
            totals[year]["raises_bonuses"] += float(row.get("yearly_difference", row.get("calculated_value", 0.0)))

        for row in self.store["time_off"]:
            year = str(record_year(row["start_date"]))
            ensure_year(year)
            totals[year]["time_off_days"] += float(row["duration_days"])

        for row in self.store["deductions"]:
            year = str(record_year(row["start_date"]))
            ensure_year(year)
            totals[year]["deductions_monthly"] += float(row.get("monthly_amount", 0.0))
            totals[year]["deductions_weekly"] += float(row.get("weekly_deduction", 0.0))
        for metrics in totals.values():
            metrics["raises_bonuses"] = round(metrics["raises_bonuses"], 2)
            metrics["time_off_days"] = round(metrics["time_off_days"], 2)
            metrics["deductions_monthly"] = round(metrics["deductions_monthly"], 2)
            metrics["deductions_weekly"] = round(metrics["deductions_weekly"], 2)
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
            tree.insert("", "end", values=(year, "Raises & Bonuses", "Total Yearly Difference", f"{metrics['raises_bonuses']:,.2f}"))
            tree.insert("", "end", values=(year, "Time Off", "Total Scheduled Days", f"{metrics['time_off_days']:,.2f}"))
            tree.insert("", "end", values=(year, "Deductions", "Total Monthly Amount", f"{metrics['deductions_monthly']:,.2f}"))
            tree.insert("", "end", values=(year, "Deductions", "Total Weekly Deduction", f"{metrics['deductions_weekly']:,.2f}"))

    def _launch_chart(self):
        """Launch the selected chart type in a separate matplotlib window."""
        if plt is None:
            messagebox.showerror("Charts", "matplotlib is required to launch charts.", parent=self.root)
            return
        if mplcursors is None and not self._hovertips_notice_shown:
            messagebox.showwarning(
                "Charts",
                "mplcursors is not installed. Charts will launch without hovertips.",
                parent=self.root,
            )
            self._hovertips_notice_shown = True
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
        messagebox.showerror("Charts", "Unsupported chart type selected.", parent=self.root)

    def _apply_light_horizontal_gridlines(self, axis):
        """Apply light horizontal gridlines for chart readability."""
        axis.set_axisbelow(True)
        axis.grid(axis="y", color="#d9d9d9", linestyle="-", linewidth=0.6, alpha=0.8)

    def _store_hover_cursor(self, figure, cursor):
        """Keep mplcursors cursor references alive for the active figure."""
        if not hasattr(figure, "_hover_cursors"):
            figure._hover_cursors = []
        figure._hover_cursors.append(cursor)

    def _format_hover_value(self, value, format_type):
        """Format hover tooltip numeric values for currency, percent, or plain numbers."""
        if format_type == "percent":
            return f"{value:,.2f}%"
        if format_type == "number":
            return f"{value:,.0f}"
        return f"${value:,.2f}"

    def _attach_bar_hovertips(self, figure, bar_container, series_name, x_labels=None, format_type="currency"):
        """Attach mplcursors hovertips to a matplotlib bar container."""
        if mplcursors is None:
            return
        patches = list(bar_container.patches)
        if not patches:
            return
        cursor = mplcursors.cursor(patches, hover=True)
        self._store_hover_cursor(figure, cursor)

        @cursor.connect("add")
        def _on_add(selection):
            patch = selection.artist
            index = None
            if patch in patches:
                index = patches.index(patch)
            elif selection.index is not None:
                index = int(round(selection.index))
            label = ""
            if x_labels is not None and index is not None:
                safe_index = int(index)
                if 0 <= safe_index < len(x_labels):
                    label = f"{x_labels[safe_index]}: "
            if hasattr(patch, "get_height"):
                value = float(patch.get_height())
            elif hasattr(patch, "get_width"):
                value = float(patch.get_width())
            else:
                value = float(selection.target[1]) if selection.target is not None else 0.0
            value_text = self._format_hover_value(value, format_type)
            selection.annotation.set_text(f"{series_name}\n{label}{value_text}")

    def _attach_line_hovertips(self, figure, line, series_name, x_labels=None, format_type="currency"):
        """Attach mplcursors hovertips to a matplotlib line chart series."""
        if mplcursors is None:
            return
        cursor = mplcursors.cursor(line, hover=True)
        self._store_hover_cursor(figure, cursor)

        @cursor.connect("add")
        def _on_add(selection):
            point_index = selection.index if selection.index is not None else 0
            safe_index = int(round(point_index))
            x_text = f"x={selection.target[0]:.2f}"
            if x_labels is not None and 0 <= safe_index < len(x_labels):
                x_text = str(x_labels[safe_index])
            y_value = float(selection.target[1])
            y_text = self._format_hover_value(y_value, format_type)
            selection.annotation.set_text(f"{series_name}\n{x_text}: {y_text}")

    def _launch_payroll_chart(self):
        """Launch a payroll dashboard with mixed subplots and KPI metrics."""
        yearly = self._calculate_yearly_totals()
        years = sorted(yearly.keys(), key=int)
        values = [yearly[year]["payroll"] for year in years]
        active_admin = [row for row in self.store["admin"] if row.get("active", True)]
        if not active_admin:
            messagebox.showinfo("Charts", "No payroll data available.", parent=self.root)
            return

        year_over_year_change = []
        previous = None
        for value in values:
            if previous is None or previous == 0:
                year_over_year_change.append(0.0)
            else:
                year_over_year_change.append(round(((value - previous) / previous) * 100, 2))
            previous = value

        per_employee_payroll = []
        for row in active_admin:
            name = normalize_text(row.get("employee_name", ""))
            if not name:
                continue
            annual_total = annualized_pay(row.get("pay_type", "Hourly"), float(row.get("pay", 0.0)))
            per_employee_payroll.append((name, annual_total))
        per_employee_payroll.sort(key=lambda item: item[1], reverse=True)
        top_entries = per_employee_payroll[:5]
        other_total = round(sum(amount for _, amount in per_employee_payroll[5:]), 2)
        if other_total > 0:
            top_entries.append(("Other", other_total))
        if not top_entries:
            top_entries = [("Unassigned", max(sum(values), 0.0))]

        fig, axes = plt.subplots(2, 2, figsize=(14, 9))
        fig.suptitle("Payroll Dashboard", fontsize=14, fontweight="bold")

        payroll_bars = axes[0, 0].bar(years, values, color="#1f77b4")
        self._apply_light_horizontal_gridlines(axes[0, 0])
        axes[0, 0].set_title("Annual Active Payroll")
        axes[0, 0].set_xlabel("Year")
        axes[0, 0].set_ylabel("Payroll")
        self._attach_bar_hovertips(fig, payroll_bars, "Annual Active Payroll", x_labels=years)

        (yoy_line,) = axes[0, 1].plot(years, year_over_year_change, marker="o", color="#ff7f0e")
        self._apply_light_horizontal_gridlines(axes[0, 1])
        axes[0, 1].axhline(0, color="#808080", linewidth=0.8, linestyle="--")
        axes[0, 1].set_title("Year-over-Year Payroll Change (%)")
        axes[0, 1].set_xlabel("Year")
        axes[0, 1].set_ylabel("Change %")
        self._attach_line_hovertips(
            fig,
            yoy_line,
            "Year-over-Year Payroll Change",
            x_labels=years,
            format_type="percent",
        )

        pie_labels = [name for name, _ in top_entries]
        pie_values = [amount for _, amount in top_entries]
        axes[1, 0].pie(pie_values, labels=pie_labels, autopct="%1.1f%%", startangle=90)
        axes[1, 0].set_title("Active Payroll Share by Employee")

        latest_year = years[-1]
        total_payroll = sum(values)
        average_payroll = round(total_payroll / len(values), 2) if values else 0.0
        latest_payroll = yearly[latest_year]["payroll"]
        axes[1, 1].axis("off")
        axes[1, 1].set_title("KPI Snapshot")
        axes[1, 1].text(0.05, 0.82, f"Active Employees: {len(active_admin)}", fontsize=11, weight="bold")
        axes[1, 1].text(0.05, 0.62, f"Latest Year ({latest_year}) Payroll: ${latest_payroll:,.2f}", fontsize=11)
        axes[1, 1].text(0.05, 0.42, f"All-Year Payroll Total: ${total_payroll:,.2f}", fontsize=11)
        axes[1, 1].text(0.05, 0.22, f"Average Payroll per Year: ${average_payroll:,.2f}", fontsize=11)

        plt.tight_layout(rect=(0, 0, 1, 0.96))
        plt.show()

    def _launch_raise_bonus_chart(self):
        """Launch a raises/bonuses dashboard with mixed subplots and KPI metrics."""
        records = self.store["raises_bonuses"]
        if not records:
            messagebox.showinfo("Charts", "No raise/bonus data available.", parent=self.root)
            return

        yearly_totals = {}
        yearly_raise_percent_totals = {}
        yearly_raise_percent_counts = {}
        change_count_totals = {"Raise": 0, "Bonus": 0}
        change_amount_totals = {"Raise": 0.0, "Bonus": 0.0}
        employee_totals = {}

        for row in records:
            year = str(record_year(row["effective_date"]))
            change_type = "Bonus" if normalize_text(row.get("change_type", "Raise")) == "Bonus" else "Raise"
            yearly_difference = float(row.get("yearly_difference", row.get("calculated_value", 0.0)) or 0.0)
            yearly_totals[year] = yearly_totals.get(year, 0.0) + yearly_difference
            change_count_totals[change_type] += 1
            change_amount_totals[change_type] += yearly_difference

            employee_name = normalize_text(row.get("employee_name", ""))
            if employee_name:
                employee_totals[employee_name] = employee_totals.get(employee_name, 0.0) + yearly_difference

            if change_type == "Raise":
                yearly_raise_percent_totals[year] = yearly_raise_percent_totals.get(year, 0.0) + float(
                    row.get("payroll_increase_percent", 0.0)
                )
                yearly_raise_percent_counts[year] = yearly_raise_percent_counts.get(year, 0) + 1

        years = sorted(yearly_totals.keys(), key=int)
        yearly_values = [round(yearly_totals[year], 2) for year in years]
        avg_raise_percents = [
            round(yearly_raise_percent_totals.get(year, 0.0) / yearly_raise_percent_counts.get(year, 1), 2)
            if yearly_raise_percent_counts.get(year, 0)
            else 0.0
            for year in years
        ]

        fig, axes = plt.subplots(2, 2, figsize=(14, 9))
        fig.suptitle("Raises & Bonuses Dashboard", fontsize=14, fontweight="bold")

        (difference_line,) = axes[0, 0].plot(years, yearly_values, marker="o", color="#2ca02c")
        self._apply_light_horizontal_gridlines(axes[0, 0])
        axes[0, 0].set_title("Yearly Total Difference")
        axes[0, 0].set_xlabel("Year")
        axes[0, 0].set_ylabel("Amount")
        self._attach_line_hovertips(fig, difference_line, "Yearly Total Difference", x_labels=years)

        count_labels = list(change_count_totals.keys())
        count_values = [change_count_totals[key] for key in count_labels]
        change_count_bars = axes[0, 1].bar(count_labels, count_values, color=["#1f77b4", "#ff9896"])
        self._apply_light_horizontal_gridlines(axes[0, 1])
        axes[0, 1].set_title("Change Counts by Type")
        axes[0, 1].set_ylabel("Count")
        self._attach_bar_hovertips(
            fig,
            change_count_bars,
            "Change Count",
            x_labels=count_labels,
            format_type="number",
        )

        pie_labels = list(change_amount_totals.keys())
        pie_values = [max(change_amount_totals[label], 0.0) for label in pie_labels]
        if sum(pie_values) <= 0:
            pie_values = [1.0, 0.0]
        axes[1, 0].pie(pie_values, labels=pie_labels, autopct="%1.1f%%", startangle=90)
        axes[1, 0].set_title("Amount Share by Change Type")

        top_employee_name = ""
        top_employee_value = 0.0
        if employee_totals:
            top_employee_name, top_employee_value = max(employee_totals.items(), key=lambda item: item[1])
        total_amount = round(sum(yearly_values), 2)
        total_events = sum(change_count_totals.values())
        average_yearly = round(total_amount / len(yearly_values), 2) if yearly_values else 0.0
        average_raise_percent = (
            round(sum(avg_raise_percents) / len([value for value in avg_raise_percents if value != 0]), 2)
            if any(value != 0 for value in avg_raise_percents)
            else 0.0
        )

        axes[1, 1].axis("off")
        axes[1, 1].set_title("KPI Snapshot")
        axes[1, 1].text(0.05, 0.85, f"Total Events: {total_events}", fontsize=11, weight="bold")
        axes[1, 1].text(0.05, 0.67, f"All-Year Difference Total: ${total_amount:,.2f}", fontsize=11)
        axes[1, 1].text(0.05, 0.49, f"Average Yearly Difference: ${average_yearly:,.2f}", fontsize=11)
        axes[1, 1].text(0.05, 0.31, f"Average Raise %: {average_raise_percent:.2f}%", fontsize=11)
        axes[1, 1].text(
            0.05,
            0.13,
            f"Top Employee Impact: {top_employee_name or 'N/A'} (${top_employee_value:,.2f})",
            fontsize=11,
        )

        raise_percent_axis = axes[0, 0].twinx()
        (avg_raise_line,) = raise_percent_axis.plot(years, avg_raise_percents, marker="s", linestyle="--", color="#9467bd")
        self._apply_light_horizontal_gridlines(raise_percent_axis)
        self._attach_line_hovertips(
            fig,
            avg_raise_line,
            "Average Raise %",
            x_labels=years,
            format_type="percent",
        )
        axes[0, 0].set_title("Yearly Total Difference (line) + Avg Raise % (dashed)")

        plt.tight_layout(rect=(0, 0, 1, 0.96))
        plt.show()

    def _launch_deductions_chart(self):
        """Launch a deductions dashboard with mixed subplots and KPI metrics."""
        records = self.store["deductions"]
        if not records:
            messagebox.showinfo("Charts", "No deductions data available.", parent=self.root)
            return

        weekly_by_type = {}
        monthly_by_type = {}
        yearly_monthly_totals = {}
        employer_contribution_total = 0.0
        unique_employees = set()

        for row in records:
            deduction_type = normalize_text(row.get("deduction_type", "")) or "Other"
            weekly_value = float(row.get("weekly_deduction", 0.0) or 0.0)
            monthly_value = float(row.get("monthly_amount", 0.0) or 0.0)
            employer_value = float(row.get("employer_contribution", 0.0) or 0.0)
            year = str(record_year(row["start_date"]))

            weekly_by_type[deduction_type] = weekly_by_type.get(deduction_type, 0.0) + weekly_value
            monthly_by_type[deduction_type] = monthly_by_type.get(deduction_type, 0.0) + monthly_value
            yearly_monthly_totals[year] = yearly_monthly_totals.get(year, 0.0) + monthly_value
            employer_contribution_total += employer_value

            employee_name = normalize_text(row.get("employee_name", ""))
            if employee_name:
                unique_employees.add(employee_name.casefold())

        deduction_labels = sorted(weekly_by_type.keys(), key=str.casefold)
        weekly_values = [round(weekly_by_type[label], 2) for label in deduction_labels]
        monthly_values = [round(monthly_by_type[label], 2) for label in deduction_labels]
        years = sorted(yearly_monthly_totals.keys(), key=int)
        yearly_values = [round(yearly_monthly_totals[year], 2) for year in years]

        fig, axes = plt.subplots(2, 2, figsize=(14, 9))
        fig.suptitle("Deductions Dashboard", fontsize=14, fontweight="bold")

        weekly_deduction_bars = axes[0, 0].bar(deduction_labels, weekly_values, color="#17becf")
        self._apply_light_horizontal_gridlines(axes[0, 0])
        axes[0, 0].set_title("Weekly Deductions by Type")
        axes[0, 0].set_ylabel("Weekly Deduction")
        axes[0, 0].tick_params(axis="x", rotation=20)
        self._attach_bar_hovertips(fig, weekly_deduction_bars, "Weekly Deductions", x_labels=deduction_labels)

        (monthly_trend_line,) = axes[0, 1].plot(years, yearly_values, marker="o", color="#8c564b")
        self._apply_light_horizontal_gridlines(axes[0, 1])
        axes[0, 1].set_title("Monthly Amount Trend by Start Year")
        axes[0, 1].set_xlabel("Year")
        axes[0, 1].set_ylabel("Total Monthly Amount")
        self._attach_line_hovertips(fig, monthly_trend_line, "Monthly Amount Trend", x_labels=years)

        pie_values = [value if value > 0 else 0 for value in monthly_values]
        if sum(pie_values) <= 0:
            pie_values = [1.0] + [0.0 for _ in deduction_labels[1:]]
        axes[1, 0].pie(pie_values, labels=deduction_labels, autopct="%1.1f%%", startangle=90)
        axes[1, 0].set_title("Monthly Amount Share by Type")

        total_monthly = round(sum(monthly_values), 2)
        total_weekly = round(sum(weekly_values), 2)
        average_weekly_per_employee = round(total_weekly / len(unique_employees), 2) if unique_employees else 0.0

        axes[1, 1].axis("off")
        axes[1, 1].set_title("KPI Snapshot")
        axes[1, 1].text(0.05, 0.85, f"Deduction Records: {len(records)}", fontsize=11, weight="bold")
        axes[1, 1].text(0.05, 0.67, f"Total Monthly Amount: ${total_monthly:,.2f}", fontsize=11)
        axes[1, 1].text(0.05, 0.49, f"Total Weekly Deduction: ${total_weekly:,.2f}", fontsize=11)
        axes[1, 1].text(0.05, 0.31, f"Employer Contribution Total: ${employer_contribution_total:,.2f}", fontsize=11)
        axes[1, 1].text(
            0.05,
            0.13,
            f"Avg Weekly/Employee: ${average_weekly_per_employee:,.2f}",
            fontsize=11,
        )

        plt.tight_layout(rect=(0, 0, 1, 0.96))
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
