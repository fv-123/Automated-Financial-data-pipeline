-- =============================
-- TOGGLE 2: Category mode
-- =============================
-- Edit these lists:
-- tickers: ('CTG','VCB')
-- categories: ('Profitability','Valuation')
-- date filter: '2021-01-01' AND '2025-09-10'
----------------------------------------------------

-- 0. Cleanup
DROP TABLE IF EXISTS tmp_fund_stripped_cat;
DROP VIEW IF EXISTS vw_price_fund_bins_cat CASCADE;
DROP VIEW IF EXISTS vw_dataset_cat CASCADE;

-- 1. Strip fundamentals to the tickers + categories you want (keep all indicators in those categories)
CREATE TEMP TABLE tmp_fund_stripped_cat AS
SELECT
    ticker,
    indicator_name,
    category,
    value,
    public_date,
    date       AS quarter_end_date,
    quarter_year
FROM fundamentals
WHERE ticker IN ('CTG','VCB')                      -- <<-- EDIT
  AND category IN ('Profitability','Valuation');    -- <<-- EDIT

CREATE INDEX IF NOT EXISTS idx_tmp_fund_cat_tk_pub ON tmp_fund_stripped_cat(ticker, public_date);

-- 2. Bin prices using the smaller stripped table (LATERAL for last public_date <= price_date)
CREATE TEMP VIEW vw_price_fund_bins_cat AS
SELECT
    p.ticker,
    p.date AS price_date,
    p.open_price,
    p.high_price,
    p.low_price,
    p.close_price,
    p.volume,
    c.industry,
    f.public_date AS fundamental_public_date
FROM prices p
JOIN companies c USING (ticker)
LEFT JOIN LATERAL (
    SELECT public_date
    FROM tmp_fund_stripped_cat f2
    WHERE f2.ticker = p.ticker
      AND f2.public_date <= p.date
    ORDER BY public_date DESC
    LIMIT 1
) f ON TRUE
WHERE p.ticker IN ('CTG','VCB');  -- <<-- keep aligned

-- 3. Join to the stripped fundamentals but DO NOT pivot in SQL (return indicator rows)
CREATE TEMP VIEW vw_dataset_cat AS
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
    tf.quarter_end_date,
    tf.quarter_year,
    pb.price_date - pb.fundamental_public_date AS days_since_public,
    tf.indicator_name,
    tf.value
FROM vw_price_fund_bins_cat pb
LEFT JOIN tmp_fund_stripped_cat tf
  ON pb.ticker = tf.ticker
 AND pb.fundamental_public_date = tf.public_date
ORDER BY pb.ticker, pb.price_date, tf.indicator_name;

-- 4. Final filtered select (date range)
SELECT *
FROM vw_dataset_cat
WHERE price_date BETWEEN '2021-01-01' AND '2025-09-10'  -- <<-- EDIT
ORDER BY ticker, price_date, indicator_name;
