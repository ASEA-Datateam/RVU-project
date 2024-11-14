# %%
# Import Packages
import pandas as pd
import statsmodels.api as sm
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import scipy as sp
import pyreadstat as prt
import statistics

from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.writer.excel import ExcelWriter
from openpyxl import load_workbook
import io


# %%
#define a function for basic statistichs

def get_column_statistics(df, column_names):
    """
    returns basic statistics for specified colums in a dataframe as a dataframe
    """
    stats_list= []

    for column_name in column_names:
        if column_name in df.columns:
            stats = {
                'Column':column_name,
                "Count": df[column_name].count(),
                "Mean": df[column_name].mean(),
                "Standard deviation": df[column_name].std(),
                "25th percentile": df[column_name].quantile(0.25),
                "Median": df[column_name].median(),
                "75th percentile": df[column_name].quantile(0.75),
                "Mode": statistics.mode(df[column_name]),
                "CPB Mode": (df[column_name].mean()/100)*79
            }
            stats_list.append(stats)
        else:
            raise ValueError(f"Column '{column_name}' does not exist in the DataFrame")
    
    stats_df = pd.DataFrame(stats_list)
    return stats_df

# %%
#Load the RVU file
df = pd.read_csv(r"F:\Desktop\Python project\Inkomen\Datafiles\rvu_compleet_def.csv")

# %%
df.info()

# %%
#Filter out 'defensie'
df = df.loc[df['CAO']!= 1597]

#Filter out zero incomes
df = df.loc[df['SBASISLOON']>0] 

# %%
#Get basic statistics for income and wealth
get_column_statistics(df,['SBASISLOON','SBASISLOON_IDX','LOON_BB','LOON_BB_IDX','LOON_ALL', 'LOON_ALL_IDX','UITWERK','UITWERK_IDX', 'UURLOON',"UURLOON_IDX",'UURLOON_BB','UURLOON_BB_IDX'])

# %%
#Write basic statistics for income and wealth to excel
get_column_statistics(df,['SBASISLOON','SBASISLOON_IDX','LOON_BB','LOON_BB_IDX','LOON_ALL', 'LOON_ALL_IDX', 'UITWERK','UITWERK_IDX', 'UURLOON',"UURLOON_IDX", 'UURLOON_BB','UURLOON_BB_IDX']).to_excel("rvu_analyse.xlsx", sheet_name='inkomen', index=False)

# %%
# count number of people who earn over 150k
len(df.loc[df["LOON_ALL"]>150000])

# %%
#Graph income distribution RVU with LOON_ALL Variable

graph_data = df.loc[df["LOON_ALL"]<150000]['LOON_ALL']
plt.figure(figsize=(18,9), dpi=400)
counts, bin_edges, _ = plt.hist(graph_data, color = "lightblue", bins=50, edgecolor = 'black', density=False)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

#sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_loon = df['LOON_ALL'].mean()
median_loon = df['LOON_ALL'].median()


plt.axvline(mean_loon, color='orange', linestyle = '--', label=f'mean: {mean_loon:.2f}')
plt.axvline(median_loon, color='black', linestyle = '--', label=f'median: {median_loon:.2f}')


plt.xlabel("Loon")
plt.ylabel("Aantal")

plt.title("Grafiek 1: Verdeling loon rvu'ers (inc. bb en ow) (tot 150k)")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

plt.savefig('g_loon_all.png', dpi=400)

# %%
#Create dataframe with histogram bins as background data
g_loonall=pd.DataFrame({
    'bin_start': bin_edges[:-1],
    'bin_end':bin_edges[1:],
    'count': counts
})

#Write dataframe with histogram bins to excel
with pd.ExcelWriter('rvu_analyse.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    g_loonall.to_excel(writer,sheet_name='inkomen', startrow=16, index=False, header=True)
    

# %%


# %%
graph_data = df.loc[df["LOON_ALL_IDX"]<150000]['LOON_ALL_IDX']
plt.figure(figsize=(18,9),dpi=600)
counts, bin_edges, _ = plt.hist(graph_data, color = "lightblue", bins=60, edgecolor = 'black', density=False)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

#sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_loon = df['LOON_ALL_IDX'].mean()
median_loon = df['LOON_ALL_IDX'].median()


plt.axvline(mean_loon, color='orange', linestyle = '--', label=f'mean: {mean_loon:.2f}')
plt.axvline(median_loon, color='black', linestyle = '--', label=f'median: {median_loon:.2f}')
plt.axvline(47971.13, color='green', linestyle = '--', label='mean 64: 47971.13')
plt.axvline(41746, color='purple', linestyle = '--', label='median 64: 41746')

plt.xlabel("Loon")
plt.ylabel("Aantal")

plt.title("Grafiek 2: Verdeling loon rvu'ers (inc. bb en ow, geindexeerd) (tot 150k)")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

plt.savefig("grafiek_loon_all_idx")

# %%
g_loonallidx=pd.DataFrame({
    'bin_start': bin_edges[:-1],
    'bin_end':bin_edges[1:],
    'count': counts
})

with pd.ExcelWriter('rvu_analyse.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    g_loonallidx.to_excel(writer,sheet_name='inkomen', startrow=70, index=False, header=True)
    

# %%


# %%
graph_data = df.loc[df["UITWERK_IDX"]<150000]['UITWERK_IDX']
plt.figure(figsize=(18,9),dpi=600)
counts, bin_edges, _ = plt.hist(graph_data, color = "lightblue", bins=60, edgecolor = 'black', density=False)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

#sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_loon = df['UITWERK_IDX'].mean()
median_loon = df['UITWERK_IDX'].median()


plt.axvline(mean_loon, color='orange', linestyle = '--', label=f'mean: {mean_loon:.2f}')
plt.axvline(median_loon, color='black', linestyle = '--', label=f'median: {median_loon:.2f}')
plt.axvline(47971.13, color='green', linestyle = '--', label='mean 64: 37562.91')
plt.axvline(41746, color='purple', linestyle = '--', label='median 64: 32001')

plt.xlabel("Loon")
plt.ylabel("Aantal")

plt.title("Grafiek 5: Verdeling loon rvu'ers (UITWERK, geindexeerd) (tot 150k)")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

plt.savefig("grafiek_UITWERK_idx")

# %%
g_loonallidx=pd.DataFrame({
    'bin_start': bin_edges[:-1],
    'bin_end':bin_edges[1:],
    'count': counts
})

with pd.ExcelWriter('rvu_analyse_update.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    g_loonallidx.to_excel(writer,sheet_name='test', startrow=1, index=False, header=True)
    

# %%
sector_counts = df['CAO'].value_counts()
large_sectors = sector_counts[sector_counts >= 100].index.tolist()
filtered_df = df[df['CAO'].isin(large_sectors)]

wage_treshold_mean = df['LOON_ALL'].mean()

result = []
for sector in filtered_df['CAO'].unique():
    sector_data = filtered_df[filtered_df["CAO"]== sector]
    above_treshold = (sector_data["LOON_ALL"]>wage_treshold_mean).sum()
    total_persons = len(sector_data)
    percentage_above_treshold = (above_treshold/total_persons)*100
    result.append({'sector':sector,
               'percentage boven gemiddelde (46568)': percentage_above_treshold})

df_percentages_boven_mean = pd.DataFrame(result)

# %%
#Create library for replacing CAO codes with sector names
sectorcode_to_text = {1659:"ING" , 1633: "ABN AMRO", 3707: "DSM", 72: "KLM Grondpersoneel", 1636: "Politie", 1646: "Rijksoverheid", 603: "NS", 637: "Verzekeringsbedrijf", 1630: "Gemeente", 1022: "MBO", 1188: "Voortgezet Onderwijs", 2881: "GVB", 1494: "Primair onderwijs", 10: "Bouw en infra", 1618: "UMC" , 9999: "Geen CAO" , 487: "Metalectro" , 163: "OV" , 21: "Goederenvervoer weg", 759: "Schilders" , 156: "Ziekenhuizen" , 2297: "Metaal en techniek: installatie" , 823: "Motorv en 2-wieler bedrijf" , 824: "Metaal en Techniek: bewerking", 51: "Timmerindustrie", 1521:"postNL" , 2948: "VVT", 49: "VVT:verpl. verz. huizen", 1345: "Sociale werkvoorziening", 750: "Contract catering"}

# %%

df_opl = df['OPLNIVSOI2021AGG4HBmetNIRWO'].value_counts().reset_index()

with pd.ExcelWriter('rvu_analyse.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    df_opl.to_excel(writer,sheet_name='overige kenmerken', startrow=1,startcol=0, index=False, header=True)

# %%
#Count cases where education level is unknown
df['OPLNIVSOI2021AGG4HBmetNIRWO'].isna().sum()

# %%
#Create and write to excel dataframe with counts for gender
df_geslacht = df['GBAGESLACHT_x'].value_counts().reset_index()

with pd.ExcelWriter('rvu_analyse.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    df_geslacht.to_excel(writer,sheet_name='overige kenmerken', startrow=1,startcol=5, index=False, header=True)

# %%
#Create and write to excel dataframe with basic statistics for hours worked by RVUers
df['SREGULIEREUREN_WEEK'] = df['SREGULIEREUREN']/52
get_column_statistics(df, ['SREGULIEREUREN_WEEK'])

with pd.ExcelWriter('rvu_analyse.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    get_column_statistics(df, ['SREGULIEREUREN_WEEK']).to_excel(writer,sheet_name='overige kenmerken', startrow=8,startcol=4, index=False, header=True)

# %%
graph_data = df.loc[(df["UURLOON_BB"]<100) & (df["UURLOON_BB"]>10)]['UURLOON_BB']
plt.figure(figsize=(18,9), dpi=400)
counts, bin_edges, _ = plt.hist(graph_data, color = "lightblue", bins=32, edgecolor = 'black', density=False)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

#sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_loon = df['UURLOON_BB'].mean()
median_loon = df['UURLOON_BB'].median()


plt.axvline(mean_loon, color='orange', linestyle = '--', label=f'mean: {mean_loon:.2f}')
plt.axvline(median_loon, color='black', linestyle = '--', label=f'median: {median_loon:.2f}')


plt.xlabel("Uurloon")
plt.ylabel("Aantal")

plt.title("Grafiek 3: Verdeling uurloon rvu'ers (inc. bb) (tot 100 euro per uur, en vanaf 10 euro per uur)")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

plt.savefig('g_loon_all_uur.png', dpi=400)

# %%
g_loonall_uur=pd.DataFrame({
    'bin_start': bin_edges[:-1],
    'bin_end':bin_edges[1:],
    'count': counts
})

with pd.ExcelWriter('rvu_analyse.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    g_loonall_uur.to_excel(writer,sheet_name='inkomen', startrow=16, startcol=16, index=False, header=True)
    

# %%
graph_data = df.loc[(df["UURLOON_BB_IDX"]<110) & (df["UURLOON_BB_IDX"]>10)]['UURLOON_BB_IDX']
plt.figure(figsize=(18,9), dpi=400)
counts, bin_edges, _ = plt.hist(graph_data, color = "lightblue", bins=32, edgecolor = 'black', density=False)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

#sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_loon = df['UURLOON_BB_IDX'].mean()
median_loon = df['UURLOON_BB_IDX'].median()


plt.axvline(mean_loon, color='orange', linestyle = '--', label=f'mean: {mean_loon:.2f}')
plt.axvline(median_loon, color='black', linestyle = '--', label=f'median: {median_loon:.2f}')
plt.axvline(34.67, color='purple', linestyle = '--', label='mean 64: 34.67 ')
plt.axvline(29.69, color='green', linestyle = '--', label='median 64: 29.69 ')


plt.xlabel("Uurloon")
plt.ylabel("Aantal")

plt.title("Grafiek 4: Verdeling uurloon rvu'ers (inc. bb, geindexeerd) (tot 110 euro per uur, en vanaf 10 euro per uur)")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

plt.savefig('g_loon_all_uur_idx.png', dpi=400)

# %%
g_loonall_uur=pd.DataFrame({
    'bin_start': bin_edges[:-1],
    'bin_end':bin_edges[1:],
    'count': counts
})

with pd.ExcelWriter('rvu_analyse.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    g_loonall_uur.to_excel(writer,sheet_name='inkomen', startrow=52, startcol=16, index=False, header=True)
    

# %%
# calculate share of population excluded by various income limits

print("grens: 39.777")
print(((len(df.loc[df["UITWERK_IDX"]>39777]))/(len(df)))*100)

print("grens: 44.000")
print(((len(df.loc[df["UITWERK_IDX"]>44000]))/(len(df)))*100)

print("grens: 73.850")
print(((len(df.loc[df["UITWERK_IDX"]>73850]))/(len(df)))*100)

print("grens: 88.000")
print(((len(df.loc[df["UITWERK_IDX"]>88000]))/(len(df)))*100)

# %%
#calculate percentafge of people in sectors with more than 100RVU cases that are excluded by a certain income limit, and write dataframe to excel

sector_counts = df['CAO'].value_counts()
large_sectors = sector_counts[sector_counts >= 100].index.tolist()
filtered_df = df[df['CAO'].isin(large_sectors)]

wage_treshold_mean = 39777
result = []
for sector in filtered_df['CAO'].unique():
    sector_data = filtered_df[filtered_df["CAO"]== sector]
    above_treshold = (sector_data["LOON_ALL_IDX"]>wage_treshold_mean).sum()
    total_persons = len(sector_data)
    percentage_above_treshold = (above_treshold/total_persons)*100
    result.append({'sector':sector,
               'percentage': percentage_above_treshold})

df_percentage_boven = pd.DataFrame(result)

df_percentage_boven['sector'] = df_percentage_boven['sector'].replace(sectorcode_to_text)
values_to_drop = ['ING','ABN AMRO', 'KLM Grondpersoneel',"DSM", "NS", "GVB", "UMC", "postNL"]
df_percentage_boven = df_percentage_boven[~df_percentage_boven['sector'].isin(values_to_drop)]

df_percentage_boven = df_percentage_boven.sort_values(by='percentage', ascending=False).reset_index()
df_percentage_boven.drop(['index'], axis=1, inplace=True)

with pd.ExcelWriter('rvu_analyse_update.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    df_percentage_boven.to_excel(writer,sheet_name='sectoren', startrow=1,startcol=0, index=False, header=True)

# %%
# Create and write to excel a dataframe with summary statistics for different asset variables
with pd.ExcelWriter('rvu_analyse_update.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    get_column_statistics(df,['VEHW1000VERH',"VEHW1121WONH", "VEHW1110FINH","VEHW1111BANH"]).to_excel(writer,sheet_name='vermogen', startrow=1,startcol=0, index=False, header=True)


# %%
# Create graph for asset distribution

graph_data = df.loc[(df['VEHW1000VERH']>-130000) & (df['VEHW1000VERH']<1000000)]['VEHW1000VERH']
plt.figure(figsize=(16,6))
counts, bin_edges, _ = plt.hist(graph_data, color = "lightblue", bins=100, edgecolor = 'black', density=False)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

#sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_verm = df['VEHW1000VERH'].mean()
median_verm = df['VEHW1000VERH'].median()
mode_verm = statistics.mode(df['VEHW1000VERH'])


plt.axvline(mean_verm, color='blue', linestyle = '--', label=f'mean: {mean_verm:.2f}')
plt.axvline(median_verm, color='purple', linestyle = '--', label=f'median: {median_verm:.2f}')
plt.axvline(mode_verm, color='green', linestyle = '--', label=f'mode: {mode_verm:.2f}')
#plt.axvline(CPB_mode_loon_64, color='orange', linestyle = '--', label=f'CPB mode: {CPB_mode_loon_64:.2f}')

plt.xlabel("totaal vermogen")
plt.ylabel("Aantal")

plt.title("Graph v_1: verdeling vrmogen (totale vermogen, min -130k, max 1.m)")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

plt.savefig('g_verm_bank_rvu.png', dpi=400)



# %%
g_data=pd.DataFrame({
    'bin_start': bin_edges[:-1],
    'bin_end':bin_edges[1:],
    'count': counts
})

print(g_data)

with pd.ExcelWriter('rvu_analyse_update.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    g_data.to_excel(writer,sheet_name='vermogen', startrow=8, index=False, header=True)
    

# %%
g_data.loc[g_data['count']<10]

# %%
# Create graph for wage distribution for people 64 years of age. 

graph_data = df.loc[(df['VEHW1111BANH']>-130000) & (df['VEHW1111BANH']<300000)]['VEHW1111BANH']
plt.figure(figsize=(16,6))
counts, bin_edges, _ = plt.hist(graph_data, color = "lightblue", bins=100, edgecolor = 'black', density=False)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

#sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_verm = df['VEHW1111BANH'].mean()
median_verm = df['VEHW1111BANH'].median()
mode_verm = statistics.mode(df['VEHW1111BANH'])


plt.axvline(mean_verm, color='blue', linestyle = '--', label=f'mean: {mean_verm:.2f}')
plt.axvline(median_verm, color='purple', linestyle = '--', label=f'median: {median_verm:.2f}')
plt.axvline(mode_verm, color='green', linestyle = '--', label=f'mode: {mode_verm:.2f}')
#plt.axvline(CPB_mode_loon_64, color='orange', linestyle = '--', label=f'CPB mode: {CPB_mode_loon_64:.2f}')

plt.xlabel("totaal vermogen")
plt.ylabel("Aantal")

plt.title("Graph v_2: verdeling vrmogen (bak- en spaartegoeden, min -130k, max 300k)")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

plt.savefig('g_vermb_rvu.png', dpi=400)



# %%
g_data=pd.DataFrame({
    'bin_start': bin_edges[:-1],
    'bin_end':bin_edges[1:],
    'count': counts
})

print(g_data)

with pd.ExcelWriter('rvu_analyse_update.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    g_data.to_excel(writer,sheet_name='vermogen', startrow=10, startcol=3,index=False, header=True)
    

# %%
g_data.loc[g_data['count']<10]

# %%
fpath_veh_22 = r"G:\InkomenBestedingen\VEHTAB\VEH2022TABV1.sav" # set file path VEHTAB 2022
metadata_veh_22 = prt.read_sav(fpath_veh_22, metadataonly=True) # load metadata VEHTAB 2022
columns_to_load_veh_22 = ['RINPERSOONHKW', 'VEHW1000VERH', 'VEHW1121WONH'] #select columns to load from VEHTAB 2022
df_veh_22 = pd.read_spss(fpath_veh_22, usecols=columns_to_load_veh_22) #load VEHTAB 2022 data 

# %%
df_veh_22['VEHW1000VERH']

# %%
pd.to_numeric(df_veh_22['VEHW1000VERH'], errors='coerce')

# %%



