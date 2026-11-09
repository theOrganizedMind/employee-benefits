# Employee Benefits Manager

This project provides a tabbed Tkinter application for managing employee
benefits and compensation data, including time off, performance reviews,
deductions, raises/bonuses, reminders, charts, and a dedicated Gantt timeline view.

## Project Structure

- **employee_benefits.py**: Main tabbed Tkinter application with CRUD,
  imports/exports, reminders, charts, and settings.
- **employee.py**: Defines the `Employee` class and its compensation and
  insurance calculations.
- **json_files/employee_details.json**: Private employee configuration used
  when it exists.
- **employee_details_public.json**: Public fallback employee configuration.
- **json_files/employee_benefits_records.json**: Persistent GUI data store for
  app records and settings (private).
- **employee_benefits_records_public.json**: Public fallback records store used
  when private records are unavailable.
- **payroll.py**: Standalone payroll totals example using its employee map.
- **tests/test_employee_benefits.py**: Pytest unit and GUI interaction tests.
- **requirements.txt**: Dependency list (includes `pytest`, `matplotlib`,
  `openpyxl`, and `tkcalendar`).
- **LICENSE.txt**: MIT license.

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/theOrganizedMind/employee-benefits.git
    cd employee-benefits
    ```

2. Install the listed packages:
    ```sh
    pip install -r requirements.txt
    ```

## Usage

### Employee Benefits GUI

Run the application from the project directory:

```sh
python employee_benefits.py
```

The GUI includes these tabs:

- **Time Off**
  - Track Vacation, Sick, and Not Paid requests
  - Automatically calculate duration from Start/End dates
  - Warn when new requests overlap existing time off for the same employee
  - Group/sort by Year and Employee Name
  - Import/export CSV and XLSX, plus downloadable templates
- **Performance**
  - Track review date, rating, summary, and notes
  - Uses a 1-5 rating dropdown
  - Import/export CSV and XLSX, plus downloadable templates
- **Deductions**
  - Track BCBS, Loan, and Child Support deductions
  - Weekly deduction formula:
    `(Monthly Amount - Employer Contribution) / 4`
  - Import/export CSV and XLSX, plus downloadable templates
- **Raises & Bonuses**
  - Track Raise and Bonus records with from/to compensation and yearly deltas
  - Auto-populates From amounts from selected employee compensation
  - Import/export CSV and XLSX, plus downloadable templates
- **Admin**
  - Manage employee hire/termination, pay type/pay, status, and notes
  - Track available vacation/sick hours and active/inactive status
  - Import/export CSV and XLSX, plus downloadable templates
- **Settings**
  - Configure preset defaults:
    - Default Vacation Hours
    - Default Sick Hours
    - Default Employer Contribution
  - Presets are persisted and used by forms and annual resets
- **Reminders**
  - Show upcoming time off and hire-date anniversaries
  - Export reminder view to CSV/XLSX
- **Charts**
  - Launch standalone matplotlib charts (outside table view)
  - Includes payroll, raises/bonuses, and deductions
  - Tracks yearly totals for payroll, raises/bonuses yearly differences, and time off
- **Gantt Chart**
  - Displays a dedicated Time Off Gantt timeline using an in-app canvas view
  - Shows one daily tick across the next 30 days starting from today
  - Includes a Time Off detail treeview below the chart

Employee name fields use autosuggest after at least 5 typed characters.
Date fields use `tkcalendar` when available, with `YYYY-MM-DD` entry fallback.

## Testing

Run the test suite with:

```sh
pytest -q
```

Tests include:

- Unit tests for core validation/calculation utilities
- GUI tests for Tkinter interactions (autosuggest, settings defaults, layout,
  deduction calculations, annual reset behavior)

If Tk cannot be initialized in your environment, GUI tests are skipped.

### Payroll Totals Script

To run the standalone payroll totals example:

```sh
python payroll.py
```

## Contributing

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Make your changes.
4. Commit your changes (`git commit -am 'Add new feature'`).
5. Push to the branch (`git push origin feature-branch`).
6. Create a new Pull Request.

## License

This project is licensed under the MIT License. 
See the LICENSE.txt file for details.
