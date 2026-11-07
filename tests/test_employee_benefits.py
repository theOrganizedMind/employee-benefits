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
        "employee_name",
        "hire_date",
        "pay_type",
        "termination_date",
        "pay",
        "available_vacation_hours",
        "notes",
        "available_sick_hours",
    ]


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
