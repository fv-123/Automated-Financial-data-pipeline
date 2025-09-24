-- =======================================================
-- FULL PIPELINE SCRIPT (Temp Views)
-- Strips fundamentals first → bins prices → joins dataset
-- =======================================================

-- =======================================================
-- 0. Drop old temp objects
-- =======================================================
DROP VIEW IF EXISTS vw_fundamentals_stripped CASCADE;
DROP VIEW IF EXISTS vw_price_fundamentals_pub_bins CASCADE;
DROP VIEW IF EXISTS vw_dataset CASCADE;

-- =======================================================
-- 1. STRIP FUNDAMENTALS
-- =======================================================
-- === Toggle 1: INDIVIDUAL INDICATORS ===
-- Example: ROE, ROA, Net profit margin
CREATE TEMP VIEW vw_fundamentals_stripped AS
SELECT *
FROM fundamentals
WHERE ticker IN ('CTG','VCB')  -- your tickers
  AND indicator_name IN ('ROE','ROA',);

-- === Toggle 2: CATEGORIES ===
-- Example: Profitability + Valuation
-- Uncomment this block instead of above if you want categories
/*
CREATE TEMP VIEW vw_fundamentals_stripped AS
SELECT *
FROM fundamentals
WHERE ticker IN ('CTG','VCB')  -- your tickers
  AND category IN ('Profitability','Valuation');
*/

-- =======================================================
-- 2. BIN PRICES TO NEAREST FUNDAMENTAL
-- =======================================================
CREATE TEMP VIEW vw_price_fundamentals_pub_bins AS
SELECT 
    p.ticker,
    p.date AS price_date,
    p.open_price,
    p.high_price,
    p.low_price,
    p.close_price,
    p.volume,
    c.industry,
    (
        SELECT MAX(f.public_date)
        FROM vw_fundamentals_stripped f
        WHERE f.ticker = p.ticker
          AND f.public_date <= p.date
    ) AS fundamental_public_date
FROM prices p
JOIN companies c ON p.ticker = c.ticker
WHERE p.ticker IN ('CTG','VCB');  -- keep tickers aligned

-- =======================================================
-- 3. JOIN WITH STRIPPED FUNDAMENTALS
-- =======================================================
CREATE TEMP VIEW vw_dataset AS
SELECT 
    pb.ticker,
    pb.price_date,
    pb.industry,
    pb.open_price,
    pb.high_price,
    pb.low_price,
    pb.close_price,
    pb.volume,
    pb.fundamental_public_date,
    f.date AS quarter_end_date,   -- original report date
    f.quarter_year,
    pb.price_date - pb.fundamental_public_date AS days_since_public,

    -- Pivot indicators (can add/remove more as needed)
    MAX(CASE WHEN f.indicator_name = 'Net profit margin' THEN f.value END) AS net_profit_margin,
    MAX(CASE WHEN f.indicator_name = 'ROA' THEN f.value END) AS roa,
    MAX(CASE WHEN f.indicator_name = 'ROE' THEN f.value END) AS roe

FROM vw_price_fundamentals_pub_bins pb
LEFT JOIN vw_fundamentals_stripped f
       ON pb.ticker = f.ticker
      AND pb.fundamental_public_date = f.public_date
GROUP BY 
    pb.ticker,
    pb.price_date,
    pb.industry,
    pb.open_price,
    pb.high_price,
    pb.low_price,
    pb.close_price,
    pb.volume,
    pb.fundamental_public_date,
    f.date,
    f.quarter_year;

-- =======================================================
-- 4. FINAL QUERY WITH DATE FILTER
-- =======================================================
SELECT *
FROM vw_dataset
WHERE price_date BETWEEN '2021-01-01' AND '2025-09-10'
ORDER BY ticker, price_date;
