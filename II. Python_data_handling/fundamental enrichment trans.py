import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
from datetime import datetime


def parse_quarter_to_date(quarter_str):
    """Convert quarter string like 'Q3/2020' to actual date"""
    try:
        quarter, year = quarter_str.split('/')
        quarter_num = int(quarter[1])  # Extract number from Q1, Q2, etc.
        year = int(year)

        # Map quarters to months (end of quarter)
        quarter_end_months = {1: 3, 2: 6, 3: 9, 4: 12}
        month = quarter_end_months[quarter_num]

        # Last day of quarter
        if month in [3, 12]:
            day = 31
        elif month == 6:
            day = 30
        else:  # September
            day = 30

        return datetime(year, month, day).strftime('%Y-%m-%d')
    except:
        return None


def categorize_indicator(indicator, statement, industry, current_category):
    """Categorize indicators based on statement type and industry"""

    indicator = str(indicator).strip()
    indicator_lower = indicator.lower()  # Convert to lowercase for case-insensitive matching

    # Banking (Finance) specific rules - ONLY for Balance Sheet and Income Statement
    if industry == "Bank" and statement in ["Balance Sheet", "Income Statement"]:
        if statement == "Balance Sheet":
            # Banking Balance Sheet Categories - case insensitive
            if "cash, gold and silver" in indicator_lower:
                return "Non-Earning Assets"
            elif "balances with the state bank" in indicator_lower:
                return "Non-Earning Assets"
            elif "placements at and loans to other credit institutions" in indicator_lower:
                return "Earning Assets"
            elif "trading securities" in indicator_lower and "provisions" not in indicator_lower:
                return "Earning Assets"
            elif "derivatives and other financial assets" in indicator_lower:
                return "Earning Assets"
            elif "loans, advances and finance leases to customers" in indicator_lower and "provisions" not in indicator_lower:
                return "Earning Assets"
            elif "debt purchased" in indicator_lower and "provision" not in indicator_lower:
                return "Earning Assets"
            elif "investment securities" in indicator_lower and "provisions" not in indicator_lower:
                return "Earning Assets"
            elif "capital contribution and other long-term investments" in indicator_lower:
                return "Earning Assets"
            elif "fixed assets" in indicator_lower:
                return "Non-Earning Assets"
            elif "investment properties" in indicator_lower:
                return "Non-Earning Assets"
            elif "other assets" in indicator_lower:
                return "Non-Earning Assets"
            elif "due to government and borrowings from the state bank" in indicator_lower:
                return "Other Liabilities"
            elif "placements and borrowings from other credit institutions" in indicator_lower:
                return "Other Liabilities"
            elif "deposits from customers" in indicator_lower:
                return "Customer Funding"
            elif "derivatives and other financial liabilities" in indicator_lower:
                return "Other Liabilities"
            elif "funds received from government" in indicator_lower or "valuable papers" in indicator_lower:
                return "Other Liabilities"
            elif "other liabilities" in indicator_lower:
                return "Other Liabilities"
            elif "capital and reserves" in indicator_lower or "minority interest" in indicator_lower:
                return "Equity"

        elif statement == "Income Statement":
            # Banking Income Statement Categories - case insensitive
            # Special rule for first 2 lines that appear before main headings
            if "interest income and similar income" in indicator_lower:
                return "Interest Income"
            elif "interest expense and similar expenses" in indicator_lower:
                return "Interest Expense"
            elif "net interest income" in indicator_lower:
                return "Net Interest"
            elif "net fee and commission income" in indicator_lower:
                return "Fee Income"
            elif "net gain/(loss) from foreign currencies" in indicator_lower or "net gain/(loss) from trading securities" in indicator_lower or "net gain/(loss) from investment securities" in indicator_lower:
                return "Trading Income"
            elif "net other income" in indicator_lower or "income from capital contribution" in indicator_lower:
                return "Other Income"
            elif "operating expenses" in indicator_lower:
                return "Operating Expenses"
            elif "provision for credit losses" in indicator_lower:
                return "Credit Provisions"
            elif "net profit" in indicator_lower or "profit before tax" in indicator_lower or "corporate income tax" in indicator_lower:
                return "Net Profit"

    # Regular company rules (applies to ALL non-Finance companies AND Finance Cash Flow/Ratios)
    else:
        if statement == "Balance Sheet":
            if "a. short-term assets" in indicator_lower:
                return "Short-term Assets"
            elif "b. long-term assets" in indicator_lower:
                return "Long-term Assets"
            elif "short-term liabilities" in indicator_lower:
                return "Short-term Liabilities"
            elif "long-term liabilities" in indicator_lower:
                return "Long-term Liabilities"
            elif "owner's equity" in indicator_lower:
                return "Owner's Equity"

        elif statement == "Income Statement":
            # Key Income Statement triggers (simplified) - case insensitive
            if "revenue" in indicator_lower and "net revenue" not in indicator_lower:
                return "Revenue"
            elif "net revenue" in indicator_lower or "cost of goods sold" in indicator_lower or "gross profit" in indicator_lower:
                return "Cost of Sales"
            elif "financial income" in indicator_lower or "financial expenses" in indicator_lower:
                return "Financial Items"
            elif "selling expenses" in indicator_lower or "general and administrative expenses" in indicator_lower:
                return "Operating Expenses"
            elif "operating profit" in indicator_lower:
                return "Operating Profit"
            elif "profit before tax" in indicator_lower or "net profit after tax" in indicator_lower:
                return "Net Profit"

        elif statement == "Cash Flow Statement":
            if "operating activities" in indicator_lower:
                return "Operating Activities"
            elif "investing activities" in indicator_lower:
                return "Investing Activities"
            elif "financing activities" in indicator_lower:
                return "Financing Activities"

        elif statement == "Ratios":
            if "valuation ratios" in indicator_lower:
                return "Valuation"
            elif "profitability ratios" in indicator_lower:
                return "Profitability"
            elif "growth rates" in indicator_lower:
                return "Growth"
            elif "liquidity ratios" in indicator_lower:
                return "Liquidity"
            elif "efficiency ratios" in indicator_lower:
                return "Efficiency"
            elif "leverage ratios" in indicator_lower:
                return "Leverage"
            elif "cashflow ratios" in indicator_lower:
                return "Cash Flow"

    # Return current category if no trigger found
    return current_category


def convert_wide_to_long():
    # Hide the root window
    root = tk.Tk()
    root.withdraw()

    # Select Excel file
    file_path = filedialog.askopenfilename(
        title="Select Excel file to convert",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not file_path:
        print("No file selected. Exiting.")
        return

    try:
        # Read Excel file
        print(f"Reading file: {file_path}")
        df = pd.read_excel(file_path, sheet_name='Master')

        print(f"Original shape: {df.shape}")
        print(f"All columns: {list(df.columns)}")

        # Reorder columns: move Ticker, Industry, Statement right after Unit
        all_columns = df.columns.tolist()

        # Extract the parts
        first_cols = all_columns[:2]  # Indicator, Unit
        last_cols = all_columns[-3:]  # Ticker, Industry, Statement
        quarterly_cols = all_columns[2:-3]  # All the quarterly date columns

        print(f"First columns: {first_cols}")
        print(f"Last columns (to move): {last_cols}")
        print(f"Quarterly columns: {quarterly_cols}")
        print(f"Number of quarterly columns: {len(quarterly_cols)}")

        # Reorder: Indicator, Unit, Ticker, Industry, Statement, then all quarterly columns
        new_column_order = first_cols + last_cols + quarterly_cols
        df_reordered = df[new_column_order]

        # Now do the melt with the reordered dataframe
        # ID variables: first 5 columns (Indicator, Unit, Ticker, Industry, Statement)
        id_vars = new_column_order[:5]

        # Value variables: all the quarterly columns (everything after the first 5)
        value_vars = new_column_order[5:]

        print(f"ID variables: {id_vars}")
        print(f"Value variables: {value_vars}")
        print(f"Number of value variables: {len(value_vars)}")

        # Convert to long format using melt
        df_long = pd.melt(df_reordered,
                          id_vars=id_vars,
                          value_vars=value_vars,
                          var_name='Period',
                          value_name='Value')

        print(f"Long format shape after melt: {df_long.shape}")

        # Split Period into Quarter, Year, and Date
        print("Splitting Period into Quarter, Year, and Date...")
        df_long['Quarter'] = df_long['Period'].str.extract(r'(Q[1-4])')
        df_long['Year'] = df_long['Period'].str.extract(r'/(\d{4})').astype('Int64')
        df_long['Date'] = df_long['Period'].apply(parse_quarter_to_date)

        # Add Category column
        print("Adding Category column...")
        df_long['Category'] = ""

        # Group by Ticker, Industry, Statement to process each group separately
        print("Categorizing indicators...")

        def categorize_group(group):
            current_category = "Uncategorized"
            categories = []

            for idx, row in group.iterrows():
                new_category = categorize_indicator(
                    row['Indicator'],
                    row['Statement'],
                    row['Industry'],
                    current_category
                )
                if new_category != current_category:
                    current_category = new_category
                categories.append(current_category)

            group['Category'] = categories
            return group

        # Apply categorization by group
        df_long = df_long.groupby(['Ticker', 'Industry', 'Statement'], group_keys=False).apply(categorize_group)

        # Reorder columns for final output
        final_columns = ['Indicator', 'Unit', 'Ticker', 'Industry', 'Statement', 'Category', 'Quarter', 'Year', 'Date',
                         'Period', 'Value']
        df_long = df_long[final_columns]

        print(f"Final shape: {df_long.shape}")
        print(f"Final columns: {list(df_long.columns)}")

        # Generate output file names
        base_name = os.path.splitext(file_path)[0]
        csv_output = f"{base_name}_long_format.csv.gz"
        xlsx_output = f"{base_name}_long_format.xlsx"

        # Save as compressed CSV
        df_long.to_csv(csv_output, index=False, compression='gzip')
        print(f"Compressed CSV saved: {csv_output}")

        # Save as Excel
        df_long.to_excel(xlsx_output, index=False, sheet_name='Long_Format')
        print(f"Excel saved: {xlsx_output}")

        print("\nFirst few rows of long format data:")
        print(df_long.head(10))

        print(f"\nSample categorization by statement:")
        for statement in df_long['Statement'].unique():
            print(f"\n{statement}:")
            cats = df_long[df_long['Statement'] == statement]['Category'].value_counts()
            print(cats.head())

        print(f"\nConversion complete!")
        print(f"Original wide format: {df.shape[0]} rows × {df.shape[1]} columns")
        print(f"New long format: {df_long.shape[0]} rows × {df_long.shape[1]} columns")

    except Exception as e:
        print(f"Error occurred: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    convert_wide_to_long()