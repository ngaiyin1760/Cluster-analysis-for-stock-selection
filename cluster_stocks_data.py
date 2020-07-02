import pandas as pd
import numpy as np

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

#load excel
key_metrics = pd.ExcelFile('key_metrics.xlsx')
financial_ratios = pd.ExcelFile('financial_ratios.xlsx')
price = pd.read_excel('price.xlsx').set_index('Date')

#List out key financial indicators
key_metrics_list = ['cashPerShare', 'currentRatio', 'debtToAssets', 'debtToEquity', 'dividendYield', 
                    'earningsYield', 'enterpriseValueOverEBITDA', 'evToFreeCashFlow', 'evToSales', 
                    'interestCoverage', 'netDebtToEBITDA', 'pbRatio', 'peRatio', 
                    'priceToSalesRatio', 'roe']

financial_ratio_list = ['cashRatio', 'debtRatio', 'cashConversionCycle', 'ebitPerRevenue', 'grossProfitMargin',
                        'netProfitMargin', 'operatingProfitMargin', 'priceToFreeCashFlowsRatio',
                        'returnOnAssets', 'returnOnEquity', 'priceEarningsRatio']

#Pre-set dataframes
key_metrics_2017 = pd.DataFrame(columns=key_metrics_list, index=ticker_list)
key_metrics_2018 = pd.DataFrame(columns=key_metrics_list, index=ticker_list)
key_metrics_2019 = pd.DataFrame(columns=key_metrics_list, index=ticker_list)

financial_ratios_2017 = pd.DataFrame(columns=financial_ratio_list, index=ticker_list)
financial_ratios_2018 = pd.DataFrame(columns=financial_ratio_list, index=ticker_list)
financial_ratios_2019 = pd.DataFrame(columns=financial_ratio_list, index=ticker_list)

#Put data into dataframe and fill as np.nan if missing value
for ticker in ticker_list:

    df_km = pd.read_excel(key_metrics, ticker, index_col=0)
    df_fr = pd.read_excel(financial_ratios, ticker, index_col=0)
    
    for key_metric in key_metrics_list:
        try:
            key_metrics_2017.at[ticker, key_metric] = df_km['2017'][key_metric]
            key_metrics_2018.at[ticker, key_metric] = df_km['2018'][key_metric]
            key_metrics_2019.at[ticker, key_metric] = df_km['2019'][key_metric]
        except:
            key_metrics_2017.at[ticker, key_metric] = np.nan
            key_metrics_2018.at[ticker, key_metric] = np.nan
            key_metrics_2019.at[ticker, key_metric] = np.nan

    for financial_ratio in financial_ratio_list:
        try:
            financial_ratios_2017.at[ticker, financial_ratio] = df_fr['2017'][financial_ratio]
            financial_ratios_2018.at[ticker, financial_ratio] = df_fr['2018'][financial_ratio]
            financial_ratios_2019.at[ticker, financial_ratio] = df_fr['2019'][financial_ratio]
        except:
            financial_ratios_2017.at[ticker, financial_ratio] = np.nan
            financial_ratios_2018.at[ticker, financial_ratio] = np.nan
            financial_ratios_2019.at[ticker, financial_ratio] = np.nan

#Split price data by years
price_2017 = price.loc["2017-01-03 00:00:00":"2017-12-29 00:00:00"] 
price_2018 = price.loc["2018-01-02 00:00:00":"2018-12-31 00:00:00"] 
price_2019 = price.loc["2019-01-02 00:00:00":"2019-12-31 00:00:00"] 
price_2020 = price.loc["2020-01-02 00:00:00":"2020-07-01 00:00:00"] 

#Calculate returns and price volatility
price_data = pd.DataFrame(columns=ticker_list)
price_data.loc['returns_2017'] = (price_2017.loc['2017-12-29 00:00:00'] / price_2017.loc['2017-01-03 00:00:00'] - 1)*100
price_data.loc['sd_2017'] = price_2017.std(axis=0)
price_data.loc['returns_2018'] = (price_2018.loc['2018-12-31 00:00:00'] / price_2018.loc['2018-01-02 00:00:00'] - 1)*100
price_data.loc['sd_2018'] = price_2018.std(axis=0)
price_data.loc['returns_2019'] = (price_2019.loc['2019-12-31 00:00:00'] / price_2019.loc['2019-01-02 00:00:00'] - 1)*100
price_data.loc['sd_2019'] = price_2019.std(axis=0)
price_data.loc['returns_2020'] = (price_2020.loc['2020-07-01 00:00:00'] / price_2020.loc['2020-01-02 00:00:00'] - 1)*100
price_data.loc['sd_2020'] = price_2020.std(axis=0)

data_2017 = pd.concat([key_metrics_2017, financial_ratios_2017], axis=1, sort=False)
data_2018 = pd.concat([key_metrics_2018, financial_ratios_2018], axis=1, sort=False)
data_2019 = pd.concat([key_metrics_2019, financial_ratios_2019], axis=1, sort=False)

data_2017['returns'] = price_data.T['returns_2017']
data_2017['volatility'] = price_data.T['sd_2017']

data_2018['returns'] = price_data.T['returns_2018']
data_2018['volatility'] = price_data.T['sd_2018']

data_2019['returns'] = price_data.T['returns_2019']
data_2019['volatility'] = price_data.T['sd_2019']

cols = data_2017.columns.tolist()
cols = cols[-2:] + cols[:-2]

data_2017 = data_2017[cols]
data_2018 = data_2018[cols]
data_2019 = data_2019[cols]

#Fill missing dividendyield as zero
data_2017['dividendYield'] = data_2017['dividendYield'].fillna(value=0)
data_2018['dividendYield'] = data_2018['dividendYield'].fillna(value=0)
data_2019['dividendYield'] = data_2019['dividendYield'].fillna(value=0)

#Output dataframes as excel files
data_2017.to_excel("data_2017.xlsx")
data_2018.to_excel("data_2018.xlsx")
data_2019.to_excel("data_2019.xlsx")








