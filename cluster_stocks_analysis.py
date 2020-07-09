import pandas as pd

# import hierarchical clustering libraries
import scipy.cluster.hierarchy as sch
from sklearn.cluster import AgglomerativeClustering

import matplotlib.pyplot as plt

#read prepared excel files into dataframe
data_2017 = pd.read_excel('data_2017.xlsx', index_col=0)
data_2018 = pd.read_excel('data_2018.xlsx', index_col=0)
data_2019 = pd.read_excel('data_2019.xlsx', index_col=0)

#Key financial indicators selected for experiment
financial_ratios = ['returns', 'volatility', 'debtToEquity',
                    'dividendYield', 'enterpriseValueOverEBITDA',
                    'netDebtToEBITDA', 'netProfitMargin',
                    'returnOnAssets', 'operatingProfitMargin']

#drop those stocks with NA values in selected financial indicators
data_2017_extracted = data_2017[financial_ratios].dropna()
data_2018_extracted = data_2018[financial_ratios].dropna()
data_2019_extracted = data_2019[financial_ratios].dropna()

#make sure consistency of the stocks group along different time window
indices_2017 = data_2017_extracted.index.tolist()
indices_2018 = data_2018_extracted.index.tolist()
indices_2019 = data_2019_extracted.index.tolist()

def intersection(lst1, lst2): 
    return list(set(lst1) & set(lst2)) 

indices = intersection(indices_2017, indices_2018)
indices = intersection(indices, indices_2019)

data_2017_extracted = data_2017_extracted.loc[indices]
data_2018_extracted = data_2018_extracted.loc[indices]
data_2019_extracted = data_2019_extracted.loc[indices]

#standardization along columns
data_2017_extracted_std=(data_2017_extracted.loc[:, financial_ratios[2:]]-data_2017_extracted.loc[:, financial_ratios[2:]].mean())/data_2017_extracted.loc[:, financial_ratios[2:]].std()
data_2018_extracted_std=(data_2018_extracted.loc[:, financial_ratios[2:]]-data_2018_extracted.loc[:, financial_ratios[2:]].mean())/data_2018_extracted.loc[:, financial_ratios[2:]].std()
data_2019_extracted_std=(data_2019_extracted.loc[:, financial_ratios[2:]]-data_2019_extracted.loc[:, financial_ratios[2:]].mean())/data_2019_extracted.loc[:, financial_ratios[2:]].std()


#===================2017=====================
# create dendrogram
dendrogram_2017 = sch.dendrogram(sch.linkage(data_2017_extracted_std, method='ward'))
plt.axhline(y=3.5, color='r', linestyle='--')

# create clusters
hc_2017 = AgglomerativeClustering(n_clusters=13, affinity = 'euclidean', linkage = 'ward')
# save clusters for chart
y_hc_2017 = hc_2017.fit_predict(data_2017_extracted_std)

#Choose the top 3 cluster with highest returns to price volatility ratio
data_2017_extracted['cluster_2017'] = pd.Series(y_hc_2017, index=data_2017_extracted.index)
data_2017_clustered_2017 = data_2017_extracted.groupby('cluster_2017').mean()
data_2017_clustered_2017['returnsTopricevol'] = data_2017_clustered_2017['returns'] / data_2017_clustered_2017['volatility']

num_of_stocks_2017 = []
for n in range(len(set(y_hc_2017))):
    num_of_stocks_2017.append(data_2017_extracted['cluster_2017'][data_2017_extracted['cluster_2017'] == n].count())

data_2017_clustered_2017['num_of_stocks'] = pd.Series(num_of_stocks_2017, index=data_2017_clustered_2017.index)

#Plot a line graph to see the characteristics of the clusters
data_2017_extracted_std['cluster_2017'] = pd.Series(y_hc_2017, index=data_2017_extracted_std.index)
data_2017_clustered_std = data_2017_extracted_std.groupby('cluster_2017').mean()
plt.plot(data_2017_clustered_std.T)
plt.xticks(rotation=90)

#match cluster formed in 2017 to 2018 data to see the performance
data_2018_extracted['cluster_2017'] = pd.Series(y_hc_2017, index=data_2018_extracted.index)
data_2018_clustered_2017 = data_2018_extracted.groupby('cluster_2017').mean()
data_2018_clustered_2017['returnsTopricevol'] = data_2018_clustered_2017['returns'] / data_2018_clustered_2017['volatility']
data_2018_clustered_2017['num_of_stocks'] = pd.Series(num_of_stocks_2017, index=data_2018_clustered_2017.index)

#Comparison with selection purely based on priceToVol ratio
data_2017_extracted['returnsTopricevol'] = data_2017_extracted['returns'] / data_2017_extracted['volatility']
data_2017_extracted.sort_values("returnsTopricevol", inplace = True, ascending=False)

data_17_clustered_17 = data_2017_clustered_2017.loc[:, ['returns', 'volatility', 'returnsTopricevol', 'num_of_stocks']]
data_17_clustered_17.sort_values("returnsTopricevol", inplace = True, ascending=False)

#count the number of stocks in the top 3 clusters
num_2017_selected_stocks = data_17_clustered_17[:3]['num_of_stocks'].sum()

#selected cluster's index
selected_cluster_2017_indices = data_17_clustered_17[:3].index

#performance of the selected cluster in 2018
perform_18_cluster_17 = data_2018_clustered_2017.loc[selected_cluster_2017_indices]
returns_cluster_18_17 = perform_18_cluster_17['returns'].dot(perform_18_cluster_17['num_of_stocks']) / perform_18_cluster_17['num_of_stocks'].sum()

#performance of the selected top stocks in 2018
top_17 = data_2017_extracted[:num_2017_selected_stocks].index
returns_top_18_17 = data_2018_extracted.loc[top_17]['returns'].mean()

#performance comparison
print('2018: ', returns_cluster_18_17, ' vs ', returns_top_18_17)

#===================2018=====================
# create dendrogram
dendrogram_2018 = sch.dendrogram(sch.linkage(data_2018_extracted_std, method='ward'))
plt.axhline(y=3.5, color='r', linestyle='--')

# create clusters
hc_2018 = AgglomerativeClustering(n_clusters=15, affinity = 'euclidean', linkage = 'ward')
# save clusters for chart
y_hc_2018 = hc_2018.fit_predict(data_2018_extracted_std)

data_2018_extracted['cluster_2018'] = pd.Series(y_hc_2018, index=data_2018_extracted.index)
data_2018_clustered_2018 = data_2018_extracted.groupby('cluster_2018').mean()
data_2018_clustered_2018 = data_2018_clustered_2018.iloc[:, :-1]
data_2018_clustered_2018['returnsTopricevol'] = data_2018_clustered_2018['returns'] / data_2018_clustered_2018['volatility']

num_of_stocks_2018 = []
for n in range(len(set(y_hc_2018))):
    num_of_stocks_2018.append(data_2018_extracted['cluster_2018'][data_2018_extracted['cluster_2018'] == n].count())

data_2018_clustered_2018['num_of_stocks'] = pd.Series(num_of_stocks_2018, index=data_2018_clustered_2018.index)

data_2018_extracted_std['cluster_2018'] = pd.Series(y_hc_2018, index=data_2018_extracted_std.index)
data_2018_clustered_std = data_2018_extracted_std.groupby('cluster_2018').mean()

plt.plot(data_2018_clustered_std.T)
plt.xticks(rotation=90)

data_2019_extracted['cluster_2018'] = pd.Series(y_hc_2018, index=data_2019_extracted.index)
data_2019_clustered_2018 = data_2019_extracted.groupby('cluster_2018').mean().iloc[:, :-1]
data_2019_clustered_2018['returnsTopricevol'] = data_2019_clustered_2018['returns'] / data_2019_clustered_2018['volatility']
data_2019_clustered_2018['num_of_stocks'] = pd.Series(num_of_stocks_2018, index=data_2019_clustered_2018.index)

#Comparison
data_2018_extracted['returnsTopricevol'] = data_2018_extracted['returns'] / data_2018_extracted['volatility']
data_2018_extracted.sort_values("returnsTopricevol", inplace = True, ascending=False)

data_18_clustered_18 = data_2018_clustered_2018.loc[:, ['returns', 'volatility', 'returnsTopricevol', 'num_of_stocks']]
data_18_clustered_18.sort_values("returnsTopricevol", inplace = True, ascending=False)

num_2018_selected_stocks = data_18_clustered_18[:3]['num_of_stocks'].sum()

selected_cluster_2018_indices = data_18_clustered_18[:3].index

perform_19_cluster_18 = data_2019_clustered_2018.loc[selected_cluster_2018_indices]

returns_cluster_19_18 = perform_19_cluster_18['returns'].dot(perform_19_cluster_18['num_of_stocks']) / perform_19_cluster_18['num_of_stocks'].sum()
top_18 = data_2018_extracted[:num_2018_selected_stocks].index
returns_top_19_18 = data_2019_extracted.loc[top_18]['returns'].mean()

print('2019: ', returns_cluster_19_18, ' vs ', returns_top_19_18)


#===================2019=====================
# create dendrogram
dendrogram_2019 = sch.dendrogram(sch.linkage(data_2019_extracted_std, method='ward'))
plt.axhline(y=3.5, color='r', linestyle='--')

# create clusters
hc_2019 = AgglomerativeClustering(n_clusters=14, affinity = 'euclidean', linkage = 'ward')
# save clusters for chart
y_hc_2019 = hc_2019.fit_predict(data_2019_extracted_std)
data_2019_extracted['cluster_2019'] = pd.Series(y_hc_2019, index=data_2019_extracted.index)
data_2019_clustered_2019 = data_2019_extracted.groupby('cluster_2019').mean()
data_2019_clustered_2019 = data_2019_clustered_2019.iloc[:, :-2]
data_2019_clustered_2019['returnsTopricevol'] = data_2019_clustered_2019['returns'] / data_2019_clustered_2019['volatility']

num_of_stocks_2019 = []
for n in range(len(set(y_hc_2019))):
    num_of_stocks_2019.append(data_2019_extracted['cluster_2019'][data_2019_extracted['cluster_2019'] == n].count())

data_2019_clustered_2019['num_of_stocks'] = pd.Series(num_of_stocks_2019, index=data_2019_clustered_2019.index)

data_2019_extracted_std['cluster_2019'] = pd.Series(y_hc_2019, index=data_2019_extracted_std.index)
data_2019_clustered_std = data_2019_extracted_std.groupby('cluster_2019').mean()

plt.plot(data_2019_clustered_std.T)
plt.xticks(rotation=90)

data_2019_clustered_2019 = data_2019_extracted.groupby('cluster_2019').mean().iloc[:, :-2]
data_2019_clustered_2019['returnsTopricevol'] = data_2019_clustered_2019['returns'] / data_2019_clustered_2019['volatility']
data_2019_clustered_2019['num_of_stocks'] = pd.Series(num_of_stocks_2019, index=data_2019_clustered_2019.index)


#===================2020=====================
price = pd.read_excel('price.xlsx').set_index('Date')
price_2020 = price.loc["2020-01-02 00:00:00":"2020-06-26 00:00:00"] 
price_2020 = price_2020.loc[:, indices]

data_2020 = pd.DataFrame(columns=indices)
data_2020.loc['returns'] = (price_2020.loc['2020-06-26 00:00:00'] / price_2020.loc['2020-01-02 00:00:00'] - 1)*100
data_2020.loc['volatility'] = price_2020.std(axis=0)
data_2020 = data_2020.T

data_2020['cluster_2017'] = pd.Series(y_hc_2017, index=data_2020.index)
data_2020['cluster_2018'] = pd.Series(y_hc_2018, index=data_2020.index)
data_2020['cluster_2019'] = pd.Series(y_hc_2019, index=data_2020.index)

data_2020_clustered_2019 = data_2020.groupby('cluster_2019').mean().iloc[:, :-2]

data_2020_clustered_2019['returnsTopricevol'] = data_2020_clustered_2019['returns'] / data_2020_clustered_2019['volatility']

data_2020_clustered_2019['num_of_stocks'] = pd.Series(num_of_stocks_2019, index=data_2020_clustered_2019.index)

#Comparison
data_2019_extracted['returnsTopricevol'] = data_2019_extracted['returns'] / data_2019_extracted['volatility']
data_2019_extracted.sort_values("returnsTopricevol", inplace = True, ascending=False)

data_19_clustered_19 = data_2019_clustered_2019.loc[:, ['returns', 'volatility', 'returnsTopricevol', 'num_of_stocks']]
data_19_clustered_19.sort_values("returnsTopricevol", inplace = True, ascending=False)
num_2019_selected_stocks = data_19_clustered_19[:3]['num_of_stocks'].sum()
selected_2019_indices = data_19_clustered_19[:3].index
perform_20_cluster_19 = data_2020_clustered_2019.loc[selected_2019_indices]
returns_cluster_20_19 = perform_20_cluster_19['returns'].dot(perform_20_cluster_19['num_of_stocks']) / perform_20_cluster_19['num_of_stocks'].sum()

top_19 = data_2019_extracted[:num_2019_selected_stocks].index
returns_top_20_19 = data_2020.loc[top_19]['returns'].mean()

print('2020: ', returns_cluster_20_19, ' vs ', returns_top_20_19)


#==================Summary====================
cluster_2017 = pd.DataFrame(index=range(len(set(y_hc_2017))))
cluster_2017['num_of_stocks'] = pd.Series(num_of_stocks_2017, index=cluster_2017.index) 
cluster_2017['returnsTopricevol_2017'] = data_2017_clustered_2017['returnsTopricevol']
cluster_2017['returnsTopricevol_2018'] = data_2018_clustered_2017['returnsTopricevol']

cluster_2018 = pd.DataFrame(index=range(len(set(y_hc_2018))))
cluster_2018['num_of_stocks'] = pd.Series(num_of_stocks_2018, index=cluster_2018.index) 
cluster_2018['returnsTopricevol_2018'] = data_2018_clustered_2018['returnsTopricevol']
cluster_2018['returnsTopricevol_2019'] = data_2019_clustered_2018['returnsTopricevol']

cluster_2019 = pd.DataFrame(index=range(len(set(y_hc_2019))))
cluster_2019['num_of_stocks'] = pd.Series(num_of_stocks_2019, index=cluster_2019.index) 
cluster_2019['returnsTopricevol_2019'] = data_2019_clustered_2019['returnsTopricevol']
cluster_2019['returnsTopricevol_2020'] = data_2020_clustered_2019['returnsTopricevol']


#save to excel
writer = pd.ExcelWriter('summary.xlsx', engine='xlsxwriter')

cluster_2017.to_excel(writer, sheet_name='Cluster2017')
cluster_2018.to_excel(writer, sheet_name='Cluster2018')
cluster_2019.to_excel(writer, sheet_name='Cluster2019')
data_2017_extracted.to_excel(writer, sheet_name='Data2017')
data_2018_extracted.to_excel(writer, sheet_name='Data2018')
data_2019_extracted.to_excel(writer, sheet_name='Data2019')
data_2020.to_excel(writer, sheet_name='Data2020')

# Close the Pandas Excel writer and output the Excel file.
writer.save()