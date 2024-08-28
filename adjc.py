import yfinance as yf
import pandas as pd

def populate(ticker: str, startDate: str, endDate: str, adjCloseCol: str):
    # Download desired ticker data
    data = yf.download(ticker, start=startDate, end=endDate, interval='1mo')
    
    # Extract the adjusted close prices for the last trading day of each month
    data_monthly = data['Adj Close']
    
    # Create a DataFrame to hold the dates and adj. close
    df = pd.DataFrame({
        'A': data_monthly.index.strftime('%Y-%m-%d'),   # Date in column A
        adjCloseCol: data_monthly                       # Adj. Close in specified column
    })
    
    # Load the existing Excel file and write the data
    with pd.ExcelWriter('stock-data2.xlsx', mode='a', if_sheet_exists='overlay') as writer:
        df[['A']].to_excel(writer, index=False, header=False, startrow=1, startcol=0)
        df[[adjCloseCol]].to_excel(writer, index=False, header=False, startrow=1, startcol=ord(adjCloseCol) - ord('A'))

# Examples to match your Excel sheet layout
populate("AAPL", '2007-01-01', '2024-08-01', 'B')
populate("MSFT", '2007-01-01', '2024-08-01', 'C')
populate("NVDA", '2007-01-01', '2024-08-01', 'D')
populate("META", '2007-01-01', '2024-08-01', 'E')
populate("NFLX", '2007-01-01', '2024-08-01', 'F')
populate("TSLA", '2007-01-01', '2024-08-01', 'G')
