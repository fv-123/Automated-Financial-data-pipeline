NOTE: Order of code execution is presented as index (I, II or 1, 2). Order doesn't matter in folders with no indexing
# AUTOMATED FINANCIAL DATA PIPELINE (PRICES AND FUNDAMENTALS)
## Project Overview
This project implements an automated **financial data pipeline** that cleans, enriches, and aggregates stock price and fundamental data for analysis.
The main goal is to to provide a **reliable, reproducible dataset** for quantitative research or portfolio analysis.

**Key highlights:**
- Handles multiple tickers and multiple fundamental statements (Balance Sheet, Income Statement, Cashflow Statement, Ratios).
- Enriches data by leveraging patterns inherent in raw data with VBA.
- Automatically bins price data to the most recent fundamental release (`public_date`).
- Supports pivoting of selected indicators (e.g., ROA, ROE) and optional grouping by categories (e.g., Profitability, Valuation).
- Optimized for performance by pre-filtering fundamentals and reducing unnecessary scans in SQL.
- 
## Pipeline Architecture
The pipeline is organized into the following stages:
1. **Data Collection**
    - Prices: from **vnstock** library in Python, imported into CSV.
    - Fundamentals: from Vietstock, imported into xlxs file.
2. **Data Cleaning & Enrichment**
    - Automates aligning statements, adding valuable columns, removing redundant rows and vertically stacking all sheets via VBA (light weight transformation).
      <img width="1919" height="785" alt="image" src="https://github.com/user-attachments/assets/65bc34d5-17df-41f2-8b7c-20861d040d9f" />
<div align = "center"> Figure 1: Each date period corresponds to one column, which can be difficult for analysis </div>
<br>
    - Migrate to Python for heavier transformation, mainly for pivoting all date period into just one column and unifying date format for time series analysis.
    <img width="1919" height="782" alt="image" src="https://github.com/user-attachments/assets/99614d5a-9f10-43b4-a9d5-9681ef51396f" />
<div align = "center"> Figure 2: Fully pivoted dates </div>
<br>
    - Map prices to the most recent fundamental public_date.
3. **Data Aggregation & Pivoting**
    - Define relationship between Prices and Fundamentals for seamless join 
    <img width="961" height="786" alt="image" src="https://github.com/user-attachments/assets/60d989da-784e-418c-af54-d2cbd5326fde" />
    <div align = "center"> Figure 3: Prices and Fundamentals ERD defined using Star Schema </div>
<br>
    - Pivot key indicators into columns for easier analysis.
    - Supports toggling between individual indicators or category-based groups.
4. **Output**
    - Temporary or permanent dataset views for downstream analysis.
    - Exportable CSV for Python analysis and plotting.
