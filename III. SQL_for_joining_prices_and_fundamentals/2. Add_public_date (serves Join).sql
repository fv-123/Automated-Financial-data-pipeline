-- 1. Add column (if it doesn’t exist yet)
ALTER TABLE fundamentals 
ADD COLUMN IF NOT EXISTS public_date DATE;

-- 2. Update public_date with industry-specific lags
UPDATE fundamentals f
SET public_date = f.date + INTERVAL '30 days'
FROM companies c
WHERE f.ticker = c.ticker
  AND c.industry = 'Bank';

UPDATE fundamentals f
SET public_date = f.date + INTERVAL '20 days'
FROM companies c
WHERE f.ticker = c.ticker
  AND c.industry = 'Metals & Mining';

UPDATE fundamentals f
SET public_date = f.date + INTERVAL '20 days'
FROM companies c
WHERE f.ticker = c.ticker
  AND c.industry = 'Information Technology';

UPDATE fundamentals f
SET public_date = f.date + INTERVAL '28 days'
FROM companies c
WHERE f.ticker = c.ticker
  AND c.industry = 'Transportation';

UPDATE fundamentals f
SET public_date = f.date + INTERVAL '29 days'
FROM companies c
WHERE f.ticker = c.ticker
  AND c.industry = 'Oil, Gas & Consumable Fuels';

-- 3. Optional: for any other industries, just set public_date = date
UPDATE fundamentals f
SET public_date = f.date
WHERE f.public_date IS NULL;


