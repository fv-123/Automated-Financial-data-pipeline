NOTE: Order of code execution is presented as index (I, II or 1, 2). Order doesn't matter in folders with no indexing
# AUTOMATED FINANCIAL DATA PIPELINE (PRICES AND FUNDAMENTALS)
## Project Overview
This project implements an automated **financial data pipeline** that cleans, enriches, and aggregates stock price and fundamental data for analysis. The pipeline processes data for **73** Vietnamese companies from **2020 to 2025**, handling multiple financial statements and automatically resolving temporal mismatches between price and fundamental data.

**Key highlights:**
- Processes multiple tickers across 5 industry sectors
-  Handles multiple tickers and multiple fundamental statements (Balance Sheet, Income Statement, Cashflow Statement, Ratios).
- Enriches data by leveraging patterns inherent in raw data with VBA.
- Automatically bins price data to the most recent fundamental release (`public_date`) instead of end of quarter date (this is super crucial, without this data, you're essentially using future information in your backtesting)
- Supports pivoting of selected indicators (e.g., ROA, ROE) and optional grouping by categories (e.g., Profitability, Valuation).
- Optimized for performance by pre-filtering fundamentals and reducing unnecessary scans in SQL.
## Pipeline Architecture
### Flowchart
<img width="720" height="2347" alt="image" src="https://github.com/user-attachments/assets/a3ba3ef0-429b-4a71-8307-a1d6b2247f4f" />

The pipeline is organized into the following stages:
### Stage 1: Data Collection

#### Price Data
- **Source:** vnstock Python library
- **Frequency:** Daily trading data
- **Coverage:** **Q3/2020 to Q2/2025** for **60** companies
- **Format:** CSV export with OHLCV data

#### Fundamental Data
- **Source:** **Vietstock web scraping**
- **Frequency:** Quarterly reporting (Q1, Q2, Q3, Q4)
- **Coverage:** **Q3/2020 to Q2/2025** with **~500** financial indicators per company
- **Format:** Excel workbooks with separate sheets per company

### Stage 2: Data Cleaning & Enrichment (VBA)

#### Pattern Recognition & Automation
- Detects statement boundaries using patterns in raw data date formatting
- Automatically categorizes indicators by statement type (Balance Sheet → "Earning Assets", Income Statement → "Net Interest", etc.)
- Removes redundant rows such as header, summary rows
- Vertically stacks **73** company sheets into unified dataset

![Raw data structure](https://github.com/user-attachments/assets/65bc34d5-17df-41f2-8b7c-20861d040d9f)
**Figure 1:** Raw data structure - each quarter as separate column


### Stage 3: Temporal Transformation (Python)

#### Date Pivot Logic
- Converts wide format (quarters as columns) to long format (single date column)
- Splits period strings (`Q3/2020`) into structured date formats (`2020-09-30`)
- Generates `public_date` using industry-specific reporting lags:
  - Banks: +30 days
  - Metals & Mining: +20 days  
  - Information Technology: +20 days
  - Transportation: +28 days
  - Oil, Gas & Consumable Fuels: +29 days

![Pivoted temporal structure](https://github.com/user-attachments/assets/99614d5a-9f10-43b4-a9d5-9681ef51396f)
**Figure 2:** Pivoted temporal structure for time-series analysis

#### Lookahead Bias Prevention

The pipeline calculates when fundamental information becomes publicly available, preventing unrealistic backtesting scenarios where future earnings data influences past trading decisions.

### Stage 4: Data Integration & Star Schema (PostgreSQL)

#### Database Design
- **Companies** table: ticker → industry mapping
- **Fundamentals** table: **~500** indicators with calculated `public_date`
- **Prices** table: **~1300** daily price records
- **Temporal binning**: Each price date mapped to most recent available fundamental

![Star schema](https://github.com/user-attachments/assets/60d989da-784e-418c-af54-d2cbd5326fde)
**Figure 3:** Star schema with natural keys (ticker, date)

#### Join Logic
```sql
-- Bin each price date to most recent fundamental
fundamental_date = MAX(f.date) WHERE f.date <= price_date
```
This ensures fundamental values stay fixed until `price_date` reaches a new `public_date`
### Stage 5: Analytical Views & Pivoting
#### Indicator Pivoting
Selected fundamental metrics pivot from rows to columns, enabling correlation analysis and model feature engineering.
<img width="1211" height="552" alt="image" src="https://github.com/user-attachments/assets/b421a2d2-c1c7-466d-8dc5-a8ca0ed6c088" />
**Figure 4:**  Individual indicators (ROE, ROA) as columns

#### Category-Based Grouping
Indicators can be grouped by analytical purpose (Valuation, Profitability, Liquidity) for thematic analysis.
<img width="1541" height="688" alt="image" src="https://github.com/user-attachments/assets/c9846613-a9d9-4799-880f-804c0a319274" />
**Figure 5:** Category-grouped indicators for systematic analysis
## Technical Specifications

### Performance Benchmarks
- **Database size:** **[256 MB]** for complete dataset
- **Query performance:** **[48 ms]** average for pivoted views

### Configurable Parameters
- Industry reporting lags (modifiable by sector)
- Indicator selection for pivoting
- Date ranges for analysis
- Ticker universe definition

## Example Usage

This chart demonstrates the pipeline's core functionality: aligning fundamental data with stock prices using calculated `public_date` to eliminate lookahead bias.

<img width="1400" height="700" alt="image" src="https://github.com/user-attachments/assets/b6c4a0cb-478a-4829-b6f1-344d9c71f750" />

**Figure 6:** CTG stock analysis showing indexed price movements vs. fundamental indicators (ROA, ROE)


#### Temporal Alignment
- **Step-function fundamentals:** ROA/ROE update only when quarterly data becomes public (~30 days after quarter-end for banks)
- **Continuous price data:** Daily market movements show real-time investor sentiment  
- **No lookahead bias:** Fundamental changes appear only after realistic reporting delays

