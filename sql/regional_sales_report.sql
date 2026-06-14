SELECT
    company_info.company_name,
    company_info.GICS_sector,
    company_info.GICS_sub_industry,
    regional_sales.gareac AS geographic_area_code,
    SUM(regional_sales.salecs) AS total_sales
FROM regional_sales
INNER JOIN company_info
    ON company_info.company_id = regional_sales.company_id
WHERE regional_sales.reporting_date
    BETWEEN CAST(? AS DATE) AND CAST(? AS DATE)
GROUP BY
    company_info.company_name,
    company_info.GICS_sector,
    company_info.GICS_sub_industry,
    regional_sales.gareac
ORDER BY
    company_info.GICS_sector,
    company_info.GICS_sub_industry,
    company_info.company_name,
    regional_sales.gareac;
