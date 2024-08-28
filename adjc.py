import yfinance as yf
import pandas as pd

def populate(ticker: str, startDate: str, endDate: str, adjCloseCol: str, row: int):
    # Download desired ticker data
    data = yf.download(ticker, start=startDate, end=endDate, interval='1mo')
    
    # Extract the adjusted close prices for the last trading day of each month
    data_monthly = data['Adj Close']
    
    # Create a DataFrame to hold the adjusted close
    df = pd.DataFrame({
        adjCloseCol: data_monthly                            # Adj. Close
    })
    
    # Write the data to the Excel file (without modifying the date column)
    with pd.ExcelWriter('stock-data-final.xlsx', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, index=False, header=False, startrow=row-1, startcol=ord(adjCloseCol) - ord('A'))

# Examples - Populate only adjusted close prices
populate("AAPL", '2007-01-01', '2024-08-01', 'B', 2)
populate("ADM", '2007-01-01', '2024-08-01', 'C', 2)
populate("AMZN", '2007-01-01', '2024-08-01', 'D', 2)
populate("AR", '2007-01-01', '2024-08-01', 'E', 2)
populate("BLK", '2007-01-01', '2024-08-01', 'F', 2)
populate("CNQ", '2007-01-01', '2024-08-01', 'G', 2)
populate("COST", '2007-01-01', '2024-08-01', 'H', 2)
populate("CRM", '2007-01-01', '2024-08-01', 'I', 2)
populate("DE", '2007-01-01', '2024-08-01', 'J', 2)
populate("DVN", '2007-01-01', '2024-08-01', 'K', 2)
populate("DXCM", '2007-01-01', '2024-08-01', 'L', 2)
populate("JPM", '2007-01-01', '2024-08-01', 'M', 2)
populate("KO", '2007-01-01', '2024-08-01', 'N', 2)
populate("LMT", '2007-01-01', '2024-08-01', 'O', 2)
populate("NCLH", '2007-01-01', '2024-08-01', 'P', 2)
populate("NKE", '2007-01-01', '2024-08-01', 'Q', 2)
populate("NVDA", '2007-01-01', '2024-08-01', 'R', 2)
populate("ORCL", '2007-01-01', '2024-08-01', 'S', 2)
populate("PLD", '2007-01-01', '2024-08-01', 'T', 2)
populate("PLTR", '2007-01-01', '2024-08-01', 'U', 2)
populate("PYPL", '2007-01-01', '2024-08-01', 'V', 2)
populate("SEDG", '2007-01-01', '2024-08-01', 'W', 2)
populate("UBS", '2007-01-01', '2024-08-01', 'X', 2)
populate("V", '2007-01-01', '2024-08-01', 'Y', 2)