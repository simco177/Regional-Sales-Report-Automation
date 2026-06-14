import argparse
import shutil
from datetime import date, datetime, timezone
from pathlib import Path

import duckdb
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

if __package__:
    from scripts.source_data import create_source_tables
    from scripts.web_report import create_html_report
else:
    from source_data import create_source_tables
    from web_report import create_html_report


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DATABASE_PATH = PROJECT_ROOT / "database" / "regional_sales.duckdb"
CREATE_TABLES_SQL = PROJECT_ROOT / "sql" / "create_staging_tables.sql"
LOAD_TABLES_SQL = PROJECT_ROOT / "sql" / "load_staging_tables.sql"
REPORT_SQL = PROJECT_ROOT / "sql" / "regional_sales_report.sql"
SITE_OUTPUT_DIRECTORY = PROJECT_ROOT / "outputs" / "site"
REPORT_OUTPUT_DIRECTORY = PROJECT_ROOT / "regional_sales_reports"
DEFAULT_EXCEL_OUTPUT = REPORT_OUTPUT_DIRECTORY / "regional_sales_report.xlsx"
DEFAULT_HTML_OUTPUT = SITE_OUTPUT_DIRECTORY / "index.html"
DEFAULT_COMPANY_DATA = (
    PROJECT_ROOT / "data" / "Company Information.xlsx"
)
DEFAULT_SALES_DATA = (
    PROJECT_ROOT / "data" / "Regional Sales.xlsx"
)


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Generate the regional sales report."
    )
    parser.add_argument(
        "start_date",
        help="Start date in YYYY-MM-DD format",
    )
    parser.add_argument(
        "end_date",
        help="End date in YYYY-MM-DD format",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_EXCEL_OUTPUT,
        help="Location of the generated Excel report",
    )
    parser.add_argument(
        "--html-output",
        type=Path,
        default=DEFAULT_HTML_OUTPUT,
        help="Location of the generated browser report",
    )
    parser.add_argument(
        "--company-data",
        type=Path,
        default=DEFAULT_COMPANY_DATA,
        help="Company source data in CSV or Excel format",
    )
    parser.add_argument(
        "--sales-data",
        type=Path,
        default=DEFAULT_SALES_DATA,
        help="Regional sales source data in CSV or Excel format",
    )
    return parser.parse_args()


def validate_dates(start_date_text, end_date_text):
    try:
        start_date = date.fromisoformat(start_date_text)
        end_date = date.fromisoformat(end_date_text)
    except ValueError as error:
        raise ValueError(
            "dates must use YYYY-MM-DD format"
        ) from error

    if start_date > end_date:
        raise ValueError(
            "start date must not be after end date"
        )

    return start_date, end_date


def prepare_database(company_data_path, sales_data_path):
    DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)

    connection = duckdb.connect(str(DATABASE_PATH))

    try:
        connection.execute(CREATE_TABLES_SQL.read_text())
        create_source_tables(
            connection,
            company_data_path,
            sales_data_path,
        )
        connection.execute(LOAD_TABLES_SQL.read_text())
    except Exception:
        connection.close()
        raise

    return connection


def retrieve_report_data(connection, start_date, end_date):
    report_query = REPORT_SQL.read_text()

    result = connection.execute(
        report_query,
        [start_date.isoformat(), end_date.isoformat()],
    )

    column_names = [
        column[0] for column in result.description
    ]
    rows = result.fetchall()

    return column_names, rows


def create_excel_report(
    output_path,
    start_date,
    end_date,
    column_names,
    rows,
    generated_at=None,
):
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    report_sheet = workbook.active
    report_sheet.title = "Regional Sales Report"

    if generated_at is None:
        generated_at = datetime.now(timezone.utc).replace(
            microsecond=0,
            tzinfo=None,
        )

    report_sheet.merge_cells("A1:E1")
    report_sheet["A1"] = "Regional Sales Report"
    report_sheet["A1"].font = Font(
        bold=True,
        size=16,
        color="FFFFFF",
    )
    report_sheet["A1"].fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    report_sheet["A3"] = "Start Date"
    report_sheet["B3"] = start_date
    report_sheet["C3"] = "End Date"
    report_sheet["D3"] = end_date
    report_sheet["A4"] = "Rows Produced"
    report_sheet["B4"] = len(rows)
    report_sheet["C4"] = "Generated At (UTC)"
    report_sheet["D4"] = generated_at

    for label_cell in ("A3", "C3", "A4", "C4"):
        report_sheet[label_cell].font = Font(bold=True)

    report_sheet["B3"].number_format = "yyyy-mm-dd"
    report_sheet["D3"].number_format = "yyyy-mm-dd"
    report_sheet["D4"].number_format = "yyyy-mm-dd hh:mm:ss"

    header_row = 7
    report_sheet.append([])
    report_sheet.append([])
    report_sheet.append(column_names)

    for row in rows:
        report_sheet.append(list(row))

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    for cell in report_sheet[header_row]:
        cell.font = Font(
            bold=True,
            color="FFFFFF",
        )
        cell.fill = header_fill

    report_sheet.freeze_panes = f"A{header_row + 1}"

    column_widths = {
        "A": 28,
        "B": 24,
        "C": 34,
        "D": 24,
        "E": 18,
    }

    for column, width in column_widths.items():
        report_sheet.column_dimensions[column].width = width

    for cell in report_sheet["E"][header_row:]:
        cell.number_format = "#,##0.000"

    if rows:
        table_reference = (
            f"A{header_row}:E{report_sheet.max_row}"
        )

        report_table = Table(
            displayName="RegionalSalesReport",
            ref=table_reference,
        )
        report_table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        report_sheet.add_table(report_table)

    workbook.save(output_path)


def main():
    arguments = parse_arguments()

    try:
        start_date, end_date = validate_dates(
            arguments.start_date,
            arguments.end_date,
        )

        connection = prepare_database(
            arguments.company_data,
            arguments.sales_data,
        )

        try:
            column_names, rows = retrieve_report_data(
                connection,
                start_date,
                end_date,
            )
        finally:
            connection.close()

        generated_at = datetime.now(timezone.utc).replace(
            microsecond=0,
            tzinfo=None,
        )

        create_excel_report(
            arguments.output,
            start_date,
            end_date,
            column_names,
            rows,
            generated_at,
        )

        browser_excel_output = (
            arguments.html_output.parent / arguments.output.name
        )
        if arguments.output.resolve() != browser_excel_output.resolve():
            browser_excel_output.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(arguments.output, browser_excel_output)

        create_html_report(
            arguments.html_output,
            start_date,
            end_date,
            column_names,
            rows,
            browser_excel_output.name,
            generated_at,
        )

        print(f"Excel report created: {arguments.output}")
        print(f"Web report created: {arguments.html_output}")
        print(f"Rows produced: {len(rows)}")

    except (ValueError, OSError, duckdb.Error) as error:
        raise SystemExit(str(error)) from error


if __name__ == "__main__":
    main()
