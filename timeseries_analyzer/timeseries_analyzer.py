from pathlib import Path

import pandas as pd
import xlwings as xw

# Path to this file
this_dir = Path(__file__).parent

def main():
    book = xw.Book.caller()
    sheet = book.sheets[0]
    tickers = sheet['A2'].expand('down').value
    parts = []  # List to collect individual DataFrames
    for ticker in tickers:
        # "usecols" allows us to only read in the Date and Adj Close
        adj_close = pd.read_csv(this_dir.parent / "csv" / f"{ticker}.csv",
                                index_col="Date", parse_dates=["Date"],
                                usecols=["Date", "Adj Close"])
        # Rename the column into the ticker symbol
        adj_close = adj_close.rename(columns={"Adj Close": ticker})
        # Append the stock's DataFrame to the parts list
        parts.append(adj_close)
        # Data alignment at work: Combine the 4 DataFrames into a single DataFrame
        adj_close = pd.concat(parts, axis=1)
        # Only use rows where we have data for all stocks
        adj_close = adj_close.dropna()
    # Clear existing data and write out DataFrame
    sheet['C2'].expand().clear_contents()
    sheet['C2'].value = adj_close