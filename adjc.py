import yfinance as yf
import pandas as pd

def populate(ticker: str, startDate: str, endDate: str, dateCol: str, adjCloseCol: str, row: int):
    # download desired ticker data
    data = yf.download(ticker, start=startDate, end=endDate)
    
    # isolate the 'Adj Close' variable and dates
    adj_close = data['Adj Close']
    dates = data.index
    
    # create a DataFrame to hold the dates and adj. close
    df = pd.DataFrame({
        dateCol: dates.strftime('%Y-%m-%d'),    # Date
        adjCloseCol: adj_close                  # Adj. Close
    })
    
    # Write the data to the Excel file
    with pd.ExcelWriter('stock-data.xlsx', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, index=False, header=False, startrow=row-1, startcol=ord(dateCol) - ord('A'))

# Examples
populate("AAPL", '2007-01-01', '2024-08-01', 'C', 'D', 4)
populate("MSFT", '2007-01-01', '2024-08-01', 'F', 'G', 4)
populate("NVDA", '2007-01-01', '2024-08-01', 'I', 'J', 4)
populate("META", '2007-01-01', '2024-08-01', 'L', 'M', 4)
populate("NFLX", '2007-01-01', '2024-08-01', 'O', 'P', 4)
populate("TSLA", '2007-01-01', '2024-08-01', 'R', 'S', 4)
