import FundamentalAnalysis as fa
import pandas as pd

api_key = "1205eba77f04c3d59478fd0a26c9ec11"

ticker_list = ['TMUSR', 'AAPL', 'MSFT', 'AMZN', 'FB',
               'GOOGL', 'GOOG', 'INTC', 'NVDA', 'ADBE',
               'PYPL', 'CSCO', 'NFLX', 'PEP', 'TSLA',
               'CMCSA', 'AMGN', 'COST', 'TMUS', 'AVGO',
               'TXN', 'CHTR', 'QCOM', 'GILD', 'SBUX',
               'INTU', 'VRTX', 'MDLZ', 'ISRG', 'FISV',
               'BKNG', 'ADP', 'REGN', 'ATVI', 'AMD',
               'JD', 'MU', 'AMAT', 'ILMN', 'ADSK',
               'CSX', 'MELI', 'LRCX', 'ADI', 'ZM',
               'BIIB', 'EA', 'KHC', 'WBA', 'LULU',
               'EBAY', 'MNST', 'DXCM', 'EXC', 'BIDU',
               'XEL', 'WDAY', 'DOCU', 'NTES', 'SPLK',
               'ORLY', 'NXPI', 'CTSH', 'ROST', 'KLAC',
               'SNPS', 'SGEN', 'ASML', 'IDXX', 'CSGP',
               'CTAS', 'VRSK', 'MAR', 'CDNS', 'PAYX',
               'ALXN', 'PCAR', 'MCHP', 'SIRI', 'ANSS',
               'VRSN', 'FAST', 'BMRN', 'XLNX', 'INCY',
               'DLTR', 'SWKS', 'ALGN', 'CERN', 'CPRT',
               'CTXS', 'TTWO', 'MXIM', 'CDW', 'CHKP',
               'TCOM', 'WDC', 'ULTA', 'EXPE', 'NTAP',
               'FOXA', 'LBTYK', 'FOX', 'LBTYA']

for ticker in ticker_list:
    
    print(ticker)
    
    key_metrics_annually = fa.key_metrics(ticker, api_key, period="annual")
    with pd.ExcelWriter('key_metrics.xlsx', engine='openpyxl', mode='a') as writer: 
        key_metrics_annually.to_excel(writer, sheet_name=ticker)
        
    financial_ratios_annually = fa.financial_ratios(ticker, api_key, period="annual")
    with pd.ExcelWriter('financial_ratios.xlsx', engine='openpyxl', mode='a') as writer: 
        financial_ratios_annually.to_excel(writer, sheet_name=ticker)
