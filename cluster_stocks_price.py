from pandas_datareader import data as pdr

import yfinance as yf
yf.pdr_override()

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
               'FOXA', 'LBTYK']

data = pdr.get_data_yahoo(ticker_list, start="2015-01-01")
price = data.loc[:, 'Close']
price.to_excel("price.xlsx")