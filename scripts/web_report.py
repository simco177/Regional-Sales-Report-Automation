from datetime import datetime, timezone
from decimal import Decimal
from html import escape
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_PATH = PROJECT_ROOT / "report_layout" / "regional_sales_report.html"


def _display_text(value):
    if value is None or str(value).strip() == "":
        return "Not Reported"
    return str(value)


def _decimal_value(value):
    if value is None:
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value))


def _format_sales(value):
    return f"{value:,.3f}"


def group_report_rows(rows):
    sectors = {}

    for row in rows:
        company_name, sector_name, industry_name, area_name, sales = row

        company_name = _display_text(company_name)
        sector_name = _display_text(sector_name)
        industry_name = _display_text(industry_name)
        area_name = _display_text(area_name)
        sales = _decimal_value(sales)

        sector = sectors.setdefault(
            sector_name,
            {
                "name": sector_name,
                "total": Decimal("0"),
                "industries": {},
            },
        )
        industry = sector["industries"].setdefault(
            industry_name,
            {
                "name": industry_name,
                "total": Decimal("0"),
                "companies": {},
            },
        )
        company = industry["companies"].setdefault(
            company_name,
            {
                "name": company_name,
                "total": Decimal("0"),
                "areas": [],
            },
        )

        sector["total"] += sales
        industry["total"] += sales
        company["total"] += sales
        company["areas"].append(
            {
                "name": area_name,
                "sales": sales,
            }
        )

    return sectors


def _render_sector_options(sectors):
    return "\n".join(
        (
            f'<option value="{escape(sector_name, quote=True)}">'
            f"{escape(sector_name)}</option>"
        )
        for sector_name in sectors
    )


def _render_area_rows(areas):
    return "\n".join(
        (
            '<tr class="area-row">'
            f'<td class="area-name">{escape(area["name"])}</td>'
            f'<td class="sales-value">{_format_sales(area["sales"])}</td>'
            "</tr>"
        )
        for area in areas
    )


def _render_company(company, sector_name, industry_name):
    search_text = " ".join(
        [
            sector_name,
            industry_name,
            company["name"],
            *(area["name"] for area in company["areas"]),
        ]
    ).lower()

    return (
        '<details class="company-group" open '
        f'data-search="{escape(search_text, quote=True)}">'
        '<summary class="company-summary">'
        f'<span class="group-name">{escape(company["name"])}</span>'
        f'<span class="group-total">{_format_sales(company["total"])}</span>'
        "</summary>"
        '<table class="area-table">'
        "<thead><tr><th>Geographic Area</th><th>Total Sales</th></tr></thead>"
        f"<tbody>{_render_area_rows(company['areas'])}</tbody>"
        "</table>"
        "</details>"
    )


def _render_industry(industry, sector_name):
    companies = "\n".join(
        _render_company(company, sector_name, industry["name"])
        for company in industry["companies"].values()
    )

    return (
        '<details class="industry-group" open>'
        '<summary class="industry-summary">'
        '<span class="group-label">Sub-Industry</span>'
        f'<span class="group-name">{escape(industry["name"])}</span>'
        f'<span class="group-total">{_format_sales(industry["total"])}</span>'
        "</summary>"
        f'<div class="industry-content">{companies}</div>'
        "</details>"
    )


def _render_sector(sector):
    industries = "\n".join(
        _render_industry(industry, sector["name"])
        for industry in sector["industries"].values()
    )

    return (
        '<details class="sector-group" open '
        f'data-sector="{escape(sector["name"], quote=True)}">'
        '<summary class="sector-summary">'
        '<span class="group-label">Sector</span>'
        f'<span class="group-name">{escape(sector["name"])}</span>'
        f'<span class="group-total">{_format_sales(sector["total"])}</span>'
        "</summary>"
        f'<div class="sector-content">{industries}</div>'
        "</details>"
    )


def _render_report_groups(sectors):
    if not sectors:
        return (
            '<div class="no-results permanent">'
            "No report rows were found for the selected period."
            "</div>"
        )

    return "\n".join(
        _render_sector(sector) for sector in sectors.values()
    )


def create_html_report(
    output_path,
    start_date,
    end_date,
    column_names,
    rows,
    excel_filename,
    generated_at=None,
):
    del column_names

    output_path.parent.mkdir(parents=True, exist_ok=True)

    if generated_at is None:
        generated_at = datetime.now(timezone.utc).replace(
            microsecond=0,
            tzinfo=None,
        )

    sectors = group_report_rows(rows)
    template = TEMPLATE_PATH.read_text(encoding="utf-8")

    replacements = {
        "{{START_DATE}}": escape(start_date.isoformat()),
        "{{END_DATE}}": escape(end_date.isoformat()),
        "{{ROW_COUNT}}": str(len(rows)),
        "{{GENERATED_AT}}": escape(
            generated_at.strftime("%Y-%m-%d %H:%M:%S UTC")
        ),
        "{{EXCEL_FILENAME}}": escape(excel_filename, quote=True),
        "{{SECTOR_OPTIONS}}": _render_sector_options(sectors),
        "{{REPORT_GROUPS}}": _render_report_groups(sectors),
    }

    for placeholder, value in replacements.items():
        template = template.replace(placeholder, value)

    output_path.write_text(template, encoding="utf-8")
