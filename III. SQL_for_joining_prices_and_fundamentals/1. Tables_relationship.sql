-- ============================================
-- DROP OLD TABLES
-- ============================================
DROP VIEW IF EXISTS vw_fundamentals_with_prices CASCADE;
DROP VIEW IF EXISTS vw_banking_summary CASCADE;
DROP VIEW IF EXISTS vw_latest_fundamentals CASCADE;
DROP TABLE IF EXISTS fact_prices CASCADE;
DROP TABLE IF EXISTS fact_fundamentals CASCADE;
DROP TABLE IF EXISTS dim_indicator CASCADE;
DROP TABLE IF EXISTS dim_time CASCADE;
DROP TABLE IF EXISTS dim_company CASCADE;
DROP TABLE IF EXISTS staging_fundamentals CASCADE;
DROP TABLE IF EXISTS staging_prices CASCADE;
DROP TABLE IF EXISTS companies CASCADE;
DROP TABLE IF EXISTS time_periods CASCADE;
DROP TABLE IF EXISTS indicators CASCADE;
DROP TABLE IF EXISTS fundamentals CASCADE;
DROP TABLE IF EXISTS prices CASCADE;

-- ============================================
-- SIMPLE TICKER-BASED SCHEMA
-- ============================================

-- Companies (just ticker + industry)
CREATE TABLE companies (
    ticker VARCHAR(20) PRIMARY KEY,
    industry VARCHAR(200) NOT NULL
);

-- Fundamentals (ticker-based, no complex FK)
CREATE TABLE fundamentals (
    ticker VARCHAR(20) NOT NULL,
    industry VARCHAR(200) NOT NULL,
    date DATE NOT NULL,
    quarter_year VARCHAR(10) NOT NULL,
    indicator_name TEXT NOT NULL,
    unit VARCHAR(100),
    statement_type VARCHAR(100) NOT NULL,
    category VARCHAR(200) NOT NULL,
    value NUMERIC(25,4),
    
    PRIMARY KEY (ticker, date, indicator_name, unit, statement_type, category),
    FOREIGN KEY (ticker) REFERENCES companies(ticker)
);

-- Prices (ticker-based, no complex FK)
CREATE TABLE prices (
    ticker VARCHAR(20) NOT NULL,
    date DATE NOT NULL,
    open_price NUMERIC(15,4),
    high_price NUMERIC(15,4),
    low_price NUMERIC(15,4),
    close_price NUMERIC(15,4),
    volume BIGINT,
    
    PRIMARY KEY (ticker, date),
    FOREIGN KEY (ticker) REFERENCES companies(ticker)
);

-- ============================================
-- STAGING AND DATA IMPORT
-- ============================================

-- Staging for fundamentals (matches your CSV exactly)
CREATE TABLE staging_fundamentals (
    indicator TEXT,
    unit VARCHAR(100),
    ticker VARCHAR(20),
    industry VARCHAR(200),
    statement VARCHAR(100),
    category VARCHAR(200),
    quarter VARCHAR(10),
    year INTEGER,
    date DATE,
    period VARCHAR(20),
    value NUMERIC(25,4)
);

-- Staging for prices
CREATE TABLE staging_prices (
    ticker VARCHAR(20),
    date DATE,
    open_price NUMERIC(15,4),
    high_price NUMERIC(15,4),
    low_price NUMERIC(15,4),
    close_price NUMERIC(15,4),
    volume BIGINT
);

-- ============================================
-- POPULATE FROM STAGING (run after CSV import)
-- ============================================

-- Load companies
INSERT INTO companies (ticker, industry)
SELECT DISTINCT ticker, industry 
FROM staging_fundamentals
WHERE ticker IS NOT NULL
ON CONFLICT (ticker) DO NOTHING;

-- Load fundamentals
INSERT INTO fundamentals (ticker, industry, date, quarter_year, indicator_name, unit, statement_type, category, value)
SELECT 
    ticker,
    industry,  -- Now properly included
    date,
    period,
    indicator,
    COALESCE(unit, 'N/A'),
    statement,
    category,
    value
FROM staging_fundamentals
WHERE ticker IS NOT NULL AND date IS NOT NULL AND indicator IS NOT NULL
ON CONFLICT DO NOTHING;

-- Load prices (run after importing prices CSV)
INSERT INTO prices (ticker, date, open_price, high_price, low_price, close_price, volume)
SELECT ticker, date, open_price, high_price, low_price, close_price, volume
FROM staging_prices
WHERE ticker IS NOT NULL AND date IS NOT NULL
ON CONFLICT DO NOTHING;

-- ============================================
-- BASIC INDEXES FOR PERFORMANCE
-- ============================================
CREATE INDEX idx_fundamentals_ticker ON fundamentals (ticker);
CREATE INDEX idx_fundamentals_date ON fundamentals (date);
CREATE INDEX idx_fundamentals_statement_category ON fundamentals (statement_type, category);
CREATE INDEX idx_prices_ticker ON prices (ticker);
CREATE INDEX idx_prices_date ON prices (date);

-- ============================================
-- SIMPLE TEST QUERIES
-- ============================================

-- Test 1: Data counts
SELECT 
    'Companies' as table_name, COUNT(*) as count FROM companies
UNION ALL
SELECT 'Fundamentals', COUNT(*) FROM fundamentals
UNION ALL
SELECT 'Prices', COUNT(*) FROM prices;

-- Test 2: Companies with both fundamentals and prices
SELECT 
    c.ticker,
    c.industry,
    COUNT(DISTINCT f.date) as fundamental_quarters,
    COUNT(DISTINCT p.date) as price_days
FROM companies c
LEFT JOIN fundamentals f ON c.ticker = f.ticker
LEFT JOIN prices p ON c.ticker = p.ticker
GROUP BY c.ticker, c.industry
ORDER BY c.ticker;

-- Test 3: Simple ticker-based join
SELECT 
    f.ticker,
    f.quarter_year,
    f.indicator_name,
    f.value as fundamental_value,
    AVG(p.close_price) as avg_price_in_quarter
FROM fundamentals f
JOIN prices p ON f.ticker = p.ticker
WHERE f.indicator_name = 'ROA'
    AND f.ticker = 'VCB'
GROUP BY f.ticker, f.quarter_year, f.indicator_name, f.value
ORDER BY f.quarter_year;

-- Test 4: Banking summary by ticker
SELECT 
    f.ticker,
    c.industry,
    f.statement_type,
    f.category,
    COUNT(*) as metric_count,
    AVG(f.value) as avg_value
FROM fundamentals f
JOIN companies c ON f.ticker = c.ticker
WHERE c.industry = 'Bank' AND f.statement_type = 'Balance Sheet'
GROUP BY f.ticker, c.industry, f.statement_type, f.category
ORDER BY f.ticker, f.category;