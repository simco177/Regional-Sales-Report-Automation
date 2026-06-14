CREATE OR REPLACE TABLE regional_sales (
    company_id VARCHAR(20) NOT NULL,
    gareac VARCHAR(100) NOT NULL,
    salecs DECIMAL(18, 3),
    reporting_date DATE NOT NULL
);

CREATE OR REPLACE TABLE company_info (
    company_id VARCHAR(20) PRIMARY KEY NOT NULL,
    company_name VARCHAR(255) NOT NULL,
    GICS_sector VARCHAR(255) NOT NULL,
    GICS_sub_industry VARCHAR(255) NOT NULL,
    headquarters_loc VARCHAR(255)
);
