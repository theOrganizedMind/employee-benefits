import tkinter as tk

import pytest

import employee_benefits as eb


@pytest.fixture
def app(tmp_path, monkeypatch):
    """Create an isolated app instance using a temporary datastore."""
    data_path = tmp_path / "employee_benefits_records.json"
    monkeypatch.setattr(eb, "DATA_STORE_PATH", str(data_path))
    monkeypatch.setattr(
        eb,
        "load_employee_details_seed",
        lambda: [
            {
                "name": "Janet Employee",
                "2026": {"salary": 78000},
            },
            {
                "name": "Jordan Example",
                "2026": {"salary": 68000},
            },
        ],
    )
    monkeypatch.setattr(eb.messagebox, "showinfo", lambda *args, **kwargs: None)
    monkeypatch.setattr(eb.messagebox, "showerror", lambda *args, **kwargs: None)
    monkeypatch.setattr(eb.messagebox, "askyesno", lambda *args, **kwargs: True)

    try:
        root = tk.Tk()
    except tk.TclError as error:
        pytest.skip(f"Tk not available in this environment: {error}")
    root.withdraw()
    instance = eb.EmployeeBenefitsApp(root)
    root.update_idletasks()
    yield instance, root
    root.destroy()


def test_make_default_store_contains_settings():
    """Default store should include persisted settings keys."""
    store = eb.make_default_store()
    assert "settings" in store
    assert store["settings"]["default_vacation_hours"] == eb.DEFAULT_VACATION_HOURS
    assert store["settings"]["default_sick_hours"] == eb.DEFAULT_SICK_HOURS
    assert store["settings"]["default_employer_contribution"] == eb.DEFAULT_EMPLOYER_CONTRIBUTION


def test_parse_date_required_raises():
    """Required date parsing should fail on empty values."""
    with pytest.raises(ValueError, match="Start Date is required"):
        eb.parse_date("", "Start Date", required=True)


def test_calculate_duration_days_inclusive():
    """Duration should be inclusive of start and end date."""
    start = eb.parse_date("2026-01-10", "Start Date", required=True)
    end = eb.parse_date("2026-01-12", "End Date", required=True)
    assert eb.calculate_duration_days(start, end) == 3


def test_settings_save_updates_defaults_and_form_defaults(app):
    """Saving settings should persist values and drive tab defaults."""
    instance, _root = app
    instance.settings_widgets["default_vacation_hours"].set("96")
    instance.settings_widgets["default_sick_hours"].set("56")
    instance.settings_widgets["default_employer_contribution"].set("360")
    instance._save_settings()

    assert instance.store["settings"]["default_vacation_hours"] == 96.0
    assert instance.store["settings"]["default_sick_hours"] == 56.0
    assert instance.store["settings"]["default_employer_contribution"] == 360.0

    instance._clear_form("admin", refresh=False)
    instance._clear_form("deductions", refresh=False)
    assert instance.tabs["admin"].form_vars["available_vacation_hours"].get() == "96.00"
    assert instance.tabs["admin"].form_vars["available_sick_hours"].get() == "56.00"
    assert instance.tabs["deductions"].form_vars["employer_contribution"].get() == "360.00"


def test_employee_autosuggest_requires_minimum_chars(app):
    """Autosuggest should only populate values after 5+ characters."""
    instance, root = app
    combo = instance.tabs["time_off"].form_widgets["employee_name"]
    var = instance.tabs["time_off"].form_vars["employee_name"]

    var.set("Jane")
    combo.event_generate("<KeyRelease>")
    root.update_idletasks()
    assert combo.cget("values") in ((), "")

    var.set("Janet")
    combo.event_generate("<KeyRelease>")
    root.update_idletasks()
    values = root.tk.splitlist(combo.cget("values"))
    assert any("Janet Employee" == value for value in values)


def test_admin_layout_field_order(app):
    """Admin form should preserve requested left-right row order."""
    instance, _root = app
    keys = list(instance.tabs["admin"].form_vars.keys())
    assert keys == [
        "employee_id",
        "employee_name",
        "title",
        "dob",
        "dl_no_state",
        "status",
        "hire_date",
        "pay_type",
        "termination_date",
        "pay",
        "available_vacation_hours",
        "available_sick_hours",
        "notes",
    ]


def test_admin_status_defaults_to_active(app):
    """Admin form should default status selector to Active."""
    instance, _root = app
    instance._clear_form("admin", refresh=False)
    assert instance.tabs["admin"].form_vars["status"].get() == "Active"


def test_admin_clear_button_clears_all_fields(app):
    """Admin Clear action should clear all filter and form values."""
    instance, _root = app
    form = instance.tabs["admin"].form_vars
    form["pay_type"].set("Hourly")
    form["available_vacation_hours"].set("80")
    form["available_sick_hours"].set("40")
    form["status"].set("Inactive")

    instance._clear_form("admin", refresh=False, apply_defaults=False)

    assert form["pay_type"].get() == ""
    assert form["available_vacation_hours"].get() == ""
    assert form["available_sick_hours"].get() == ""
    assert form["status"].get() == ""


def test_admin_title_field_is_editable_combo(app):
    """Admin Title field should be an editable combobox."""
    instance, _root = app
    widget = instance.tabs["admin"].form_widgets["title"]
    assert widget.winfo_class() == "TCombobox"
    assert str(widget.cget("state")) == "normal"


def test_admin_treeview_column_order(app):
    """Admin tree columns should match the requested output order."""
    instance, _root = app
    assert instance.tabs["admin"].columns == [
        "Employee ID",
        "Employee Name",
        "Title",
        "DOB",
        "DL No. - State",
        "Hire Date",
        "Pay Type",
        "Pay",
        "Available Vacation (Hours)",
        "Available Sick (Hours)",
        "Status",
        "Termination Date",
        "Notes",
    ]


def test_admin_employee_id_auto_increments(app):
    """Admin Employee ID should auto-fill the next sequential value."""
    instance, _root = app
    form = instance.tabs["admin"].form_vars
    form["employee_id"].set("")
    form["employee_name"].set("Taylor Example")
    form["title"].set("Driver")
    form["dob"].set("1990-05-05")
    form["dl_no_state"].set("D123-CA")
    form["hire_date"].set("2026-01-01")
    form["termination_date"].set("")
    form["pay_type"].set("Hourly")
    form["pay"].set("30")
    form["available_vacation_hours"].set("80")
    form["available_sick_hours"].set("40")
    form["status"].set("Active")
    form["notes"].set("")

    record = instance._build_record_from_form("admin")
    assert record["employee_id"] == "003"


def test_driver_list_export_filters_active_drivers(app, monkeypatch, tmp_path):
    """Driver List export should include only active employees with Driver title."""
    instance, _root = app
    instance.store["admin"] = [
        {
            "id": "1",
            "employee_id": "001",
            "employee_name": "Active Driver",
            "title": "Driver",
            "dob": "1990-01-01",
            "dl_no_state": "A111-TX",
            "hire_date": "2020-01-01",
            "termination_date": "",
            "pay_type": "Hourly",
            "pay": 30.0,
            "available_vacation_hours": 80.0,
            "available_sick_hours": 40.0,
            "notes": "",
            "active": True,
        },
        {
            "id": "2",
            "employee_id": "002",
            "employee_name": "Inactive Driver",
            "title": "Driver",
            "dob": "1991-01-01",
            "dl_no_state": "B222-TX",
            "hire_date": "2020-02-01",
            "termination_date": "",
            "pay_type": "Hourly",
            "pay": 30.0,
            "available_vacation_hours": 80.0,
            "available_sick_hours": 40.0,
            "notes": "",
            "active": False,
        },
        {
            "id": "3",
            "employee_id": "003",
            "employee_name": "Active Non Driver",
            "title": "Manager",
            "dob": "1992-01-01",
            "dl_no_state": "C333-TX",
            "hire_date": "2020-03-01",
            "termination_date": "",
            "pay_type": "Hourly",
            "pay": 30.0,
            "available_vacation_hours": 80.0,
            "available_sick_hours": 40.0,
            "notes": "",
            "active": True,
        },
    ]
    export_path = tmp_path / "drivers.csv"
    monkeypatch.setattr(eb.filedialog, "asksaveasfilename", lambda **kwargs: str(export_path))
    monkeypatch.setattr(eb.messagebox, "showinfo", lambda *args, **kwargs: None)
    instance._export_driver_list()

    with export_path.open(encoding="utf-8") as file:
        lines = [line.strip() for line in file.readlines()]
    assert lines[0] == "Employee ID,Employee Name,Hire Date,Title,DOB,DL No. - State"
    assert any("Active Driver" in line for line in lines[1:])
    assert not any("Inactive Driver" in line for line in lines[1:])
    assert not any("Active Non Driver" in line for line in lines[1:])


def test_deductions_weekly_uses_monthly_minus_employer(app):
    """Deductions weekly value should use monthly minus employer contribution."""
    instance, _root = app
    form = instance.tabs["deductions"].form_vars
    form["employee_name"].set("Janet Employee")
    form["deduction_type"].set("BCBS")
    form["monthly_amount"].set("500")
    form["employer_contribution"].set("300")
    form["start_date"].set("2026-01-01")
    form["notes"].set("health")

    record = instance._build_record_from_form("deductions")
    assert record["weekly_deduction"] == 50.0


def test_time_off_display_includes_year_first_column(app):
    """Time off rows should display year in the first treeview column."""
    instance, _root = app
    record = {
        "id": "abc",
        "employee_name": "Janet Employee",
        "time_off_type": "Vacation",
        "start_date": "2026-06-01",
        "end_date": "2026-06-03",
        "duration_days": 3,
        "notes": "",
    }
    displayed = instance._display_row("time_off", record)
    assert displayed[0] == "2026"
    assert displayed[1] == "Janet Employee"


def test_annual_reset_uses_settings_values(app):
    """Annual reset should use configured setting values."""
    instance, _root = app
    instance.store["settings"]["default_vacation_hours"] = 88.0
    instance.store["settings"]["default_sick_hours"] = 44.0
    instance.store["meta"]["last_annual_reset_year"] = "2000"
    instance._apply_annual_reset_if_needed()

    for row in instance.store["admin"]:
        assert row["available_vacation_hours"] == 88.0
        assert row["available_sick_hours"] == 44.0


def test_hourly_compensation_snapshot_normalizes_legacy_annual_value(app):
    """Hourly snapshot should treat legacy annualized values as hourly-derived pay."""
    instance, _root = app
    admin_record = instance.store["admin"][0]
    admin_record["pay_type"] = "Hourly"
    admin_record["pay"] = 85800.0

    hourly_from, yearly_from = instance._compensation_snapshot(admin_record)
    assert hourly_from == 30.0
    assert yearly_from == 85800.0


def test_raises_form_hourly_to_syncs_yearly_to(app):
    """Hourly To Amount should drive Yearly To Amount via Employee constants."""
    instance, root = app
    admin_record = instance.store["admin"][0]
    admin_record["pay_type"] = "Hourly"
    admin_record["pay"] = 30.0
    employee_name = admin_record["employee_name"]

    form = instance.tabs["raises_bonuses"].form_vars
    form["employee_name"].set(employee_name)
    root.update_idletasks()
    assert form["hourly_from_amount"].get() == "30.00"
    assert form["yearly_from_amount"].get() == "85800.00"

    form["hourly_to_amount"].set("31")
    root.update_idletasks()
    assert form["yearly_to_amount"].get() == "88660.00"


def test_migrate_preserves_existing_raise_from_to_values(app):
    """Schema migration should not overwrite existing persisted raise from/to values."""
    instance, _root = app
    admin_row = instance.store["admin"][0]
    admin_row["pay_type"] = "Hourly"
    admin_row["pay"] = 30.0
    instance.store["raises_bonuses"] = [
        {
            "id": "r1",
            "employee_name": admin_row["employee_name"],
            "change_type": "Raise",
            "effective_date": "2026-03-01",
            "hourly_from_amount": 28.0,
            "hourly_to_amount": 30.0,
            "yearly_from_amount": 80080.0,
            "yearly_to_amount": 85800.0,
            "yearly_difference": 5720.0,
            "payroll_increase_percent": 7.1429,
            "bonus_amount": 0.0,
            "bonus_percent": 0.0,
            "notes": "",
        }
    ]

    instance._migrate_store(instance.store)
    raise_row = instance.store["raises_bonuses"][0]
    assert raise_row["hourly_from_amount"] == 28.0
    assert raise_row["hourly_to_amount"] == 30.0
    assert raise_row["yearly_difference"] == 5720.0


def test_add_due_raise_updates_admin_pay(app):
    """Adding a due raise should automatically update Admin pay."""
    instance, _root = app
    admin_row = instance.store["admin"][0]
    admin_row["pay_type"] = "Hourly"
    admin_row["pay"] = 28.0
    employee_name = admin_row["employee_name"]

    form = instance.tabs["raises_bonuses"].form_vars
    form["employee_name"].set(employee_name)
    form["change_type"].set("Raise")
    form["effective_date"].set(eb.date.today().strftime("%Y-%m-%d"))
    form["hourly_to_amount"].set("30")
    form["yearly_to_amount"].set("")
    form["bonus_amount"].set("0")
    form["bonus_percent"].set("0")
    form["notes"].set("")

    instance._add_record("raises_bonuses")
    updated_admin = instance._get_admin_record(employee_name)
    assert updated_admin["pay"] == 30.0
