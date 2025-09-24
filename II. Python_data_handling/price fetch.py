from vnstock import Vnstock
import pandas as pd
import logging

# --- Configuration ---

# Configure logging to see progress and potential issues
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Define the list of tickers you want to fetch.
# The industry information is managed in your database, not here.
tickers = [
    "PVT","PSH","BAB","CMG","LGC","VJC","HLC","PVP","SSB","STG","VCA","STB","DVP","DTL","MSB","VIB","HVN","MBB","PHP","SCS","VCB","CDN","SHB","TCB","BCA","HMC","HDB","OCB","TVD","HSG","CTG","CQN","TMG","DLG","MSR","VGS","PLX","AST","VPB","PIA","VIP","TVN","VFC","ACV","FPT","ASP","ELC","BSR","LPB","VTV","OIL","COM","TLH","POM","VTE","PJC","MVB","VOS","BMC","NAB"

]

start_date = "2020-01-01"
end_date = "2025-09-11"
output_csv_path = "prices.csv"

# --- Data Fetching ---

# 1. Instantiate Vnstock() only ONCE for efficiency.
stock_fetcher = Vnstock()
all_prices_df = []

logging.info(f"Starting to fetch data for {len(tickers)} tickers...")

# 2. Loop through each ticker to get its price history.
for ticker in tickers:
    try:
        # Use the single stock_fetcher instance to get the data
        df = stock_fetcher.stock(symbol=ticker, source='VCI').quote.history(start=start_date, end=end_date)

        if df is not None and not df.empty:
            # Add the ticker column so we know which company the price belongs to
            df["ticker"] = ticker
            all_prices_df.append(df)
            logging.info(f"Successfully fetched {len(df)} records for {ticker}")
        else:
            logging.warning(f"No data returned for ticker: {ticker}")

    except Exception as e:
        logging.error(f"An error occurred while fetching data for {ticker}: {e}")

# 3. Concatenate all the individual DataFrames into one large DataFrame.
if all_prices_df:
    prices = pd.concat(all_prices_df, ignore_index=True)

    # 4. Clean and standardize the column names to match the database table.
    prices.rename(columns={
        'time': 'date',
        'open': 'open',
        'high': 'high',
        'low': 'low',
        'close': 'close',
        'volume': 'volume',
        'ticker': 'ticker'
    }, inplace=True)

    # Ensure all column names are lowercase
    prices.columns = prices.columns.str.lower()

    # Select only the columns that exist in your `prices` table
    # New, correct order to match the database
    final_columns = ['ticker', 'date', 'open', 'high', 'low', 'close', 'volume']
    prices = prices[final_columns]

    # 5. Save the clean data to a CSV file, ready for import.
    prices.to_csv(output_csv_path, index=False)
    logging.info(f"Successfully saved all price data to {output_csv_path}")
else:
    logging.warning("No data was fetched. The output CSV file was not created.")

