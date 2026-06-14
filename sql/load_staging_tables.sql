BEGIN TRANSACTION;

INSERT INTO company_info (
    company_id,
    company_name,
    GICS_sector,
    GICS_sub_industry,
    headquarters_loc
)
SELECT
    TRIM(company_id),
    TRIM(company_name),
    TRIM(GICS_sector),
    TRIM(GICS_sub_industry),
    NULLIF(TRIM(headquarters_loc), '')
FROM company_source;

INSERT INTO regional_sales (
    company_id,
    gareac,
    salecs,
    reporting_date
)
SELECT
    TRIM(company_id),
    TRIM(gareac),
    CAST(NULLIF(TRIM(salecs), '') AS DECIMAL(18, 3)),
    CAST(NULLIF(TRIM(reporting_date), '') AS DATE)
FROM sales_source;

COMMIT;
