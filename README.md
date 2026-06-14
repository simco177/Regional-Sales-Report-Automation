# Regional Sales Report Automation

**[View Regional Sales Report](https://simco177.github.io/Regional-Sales-Report-Automation/)**

**[Download Excel Report](regional_sales_reports/regional_sales_report.xlsx?raw=1)**

## Overview

This project automates a recurring regional sales report. It combines synthetic regional sales data with company information, filters the records for a selected reporting period, and calculates total sales by company and geographic area.

The report classifies each company by GICS sector and sub-industry. The source data contains 500 fictional companies across all 11 sectors.

## Report Output

The report calculates total sales for a selected reporting period and presents the results by GICS sector, GICS sub-industry, company, and geographic area. It includes sales subtotals, sector filtering, collapsible groups, and a search function.

## Data Sources

### Regional Sales Data

`data/Regional Sales.xlsx` contains aggregate sales for each company by geographic area and reporting date.

| Field | Description |
| --- | --- |
| `company_id` | Company identifier used to join the datasets |
| `gareac` | Geographic area name |
| `salecs` | Sales recorded for the company, area, and reporting date |
| `reporting_date` | Reporting date used to select the reporting period |

### Company Information

`data/Company Information.xlsx` contains one row per company.

| Field | Description |
| --- | --- |
| `company_id` | Unique company identifier |
| `company_name` | Company name |
| `GICS_sector` | GICS sector |
| `GICS_sub_industry` | GICS sub-industry |
| `headquarters_loc` | Headquarters location |

## Reporting Workflow

```text
Regional sales and company Excel files
            |
            v
Full refresh of DuckDB staging tables
            |
            v
SQL join, date filter, grouping, and aggregation
            |
            v
Published regional sales report
```

Each run:

1. Creates the database tables when they do not already exist.
2. Clears the existing staging-table data.
3. Reloads both source Excel files in one database transaction.
4. Filters sales using the supplied start and end dates.
5. Joins sales to the company information by `company_id`.
6. Calculates total sales by company, GICS classification, and geographic area.
7. Publishes the grouped regional sales report.

The source Excel files remain the reporting inputs. The generated DuckDB file is a local processing artifact and is not committed to the repository.

## Run Locally

Create and activate a virtual environment, then install the required packages:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Generate a report by supplying the reporting period:

```bash
python scripts/generate_report.py START_DATE END_DATE
```

Dates must use `YYYY-MM-DD` format. The Excel report is saved to:

```text
regional_sales_reports/regional_sales_report.xlsx
```

The browser report is generated at `outputs/site/index.html` for publication.

## Automated Run

The `Generate Regional Sales Report` GitHub Actions workflow accepts a start date and end date. It installs the required packages, runs the automated tests, generates the report, and deploys it through GitHub Pages.

The workflow is run from the repository's **Actions** tab by supplying the start and end dates for the reporting period.

Each workflow run publishes the report for the selected period and updates the Excel workbook in `regional_sales_reports/`.

## Updating Reporting Data

The loader reads these source paths by default:

```text
data/Regional Sales.xlsx
data/Company Information.xlsx
```

Before generating a report for a new period:

1. Update the regional sales file so it contains the required reporting records and preserves the existing column structure.
2. Refresh the company information file when company identifiers or classifications change.
3. Run the report using dates covered by the updated sales data.

Because the staging tables are fully refreshed on every run, rerunning the process does not append duplicate database records.

## Project Structure

```text
Regional-Sales-Report-Automation/
├── .github/
│   └── workflows/
│       └── generate_report.yml
├── data/
│   ├── Company Information.xlsx
│   └── Regional Sales.xlsx
├── scripts/
│   ├── generate_report.py
│   ├── source_data.py
│   └── web_report.py
├── sql/
│   ├── create_staging_tables.sql
│   ├── load_staging_tables.sql
│   └── regional_sales_report.sql
├── report_layout/
│   └── regional_sales_report.html
├── regional_sales_reports/
│   └── regional_sales_report.xlsx
├── tests/
│   └── test_generate_report.py
├── .gitignore
├── requirements.txt
└── README.md
```

## Tests

Run the automated tests with:

```bash
python -m unittest tests/test_generate_report.py
```

The tests verify date validation, grouped report output, subtotal calculations, HTML escaping, empty results, and GitHub Pages workflow configuration.

## Data Notes

- `salecs` represents synthetic sales values and is stored to three decimal places.
- `reporting_date` contains quarter-end dates.
- `gareac` contains one of eight full geographic area names.
- Companies without a matching `company_id` in the company information file are excluded from the report.
