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
#load SPOLISBUS
fpath_spolis = r"G:\Spolis\SPOLISBUS\2023\SPOLISBUS2023V4.sav" # set file path spolisbus 2023
metadata_spolis = prt.read_sav(fpath_spolis, metadataonly=True) # load metadata spolisbus 2023
columns_to_load_spolisbus = ['RINPERSOON', 'SBASISLOON','SLNOWRK','SBASISUREN', 'SREGULIEREUREN', 'SBIJZONDEREBELONING'] #select columns to load from spolisbus 2023
df_spolis_23 = pd.read_spss(fpath_spolis, usecols=columns_to_load_spolisbus) #load spolisbus data 2023

df_spolis_23['LOON_BB'] = df_spolis_23['SBASISLOON'] + df_spolis_23['SBIJZONDEREBELONING']
df_spolis_23['LOON_OW'] = df_spolis_23['SBASISLOON'] + df_spolis_23['SLNOWRK']
df_spolis_23['LOON_ALL'] = df_spolis_23['SBASISLOON'] + df_spolis_23['SBIJZONDEREBELONING'] + df_spolis_23['SLNOWRK']


# %%
#Load GBAPERSOONKTAB 2023

fpath_gba_23 = "G:\Bevolking\GBAPERSOONKTAB\GBAPERSOONKTAB2023V1.sav" # set file path GBAPERSOONKTAB 2023
df_gba_23 = pd.read_spss(fpath_gba_23) # load GBAPERSOONKTAB data 

# Create age column GBA 2023
df_gba_23['GBAGEBOORTEJAAR'] = pd.to_numeric(df_gba_23['GBAGEBOORTEJAAR'], errors='coerce') # convert to numeric
df_gba_23['leeftijd'] = 2024 - df_gba_23['GBAGEBOORTEJAAR'] # calculate age and create column


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
# For each unique ID number all registered wages are summed. 

df_spolis_23_sum = df_spolis_23.groupby('RINPERSOON')[['SBASISLOON','SBASISUREN','LOON_BB', 'LOON_OW', 'LOON_ALL','SREGULIEREUREN']].sum().reset_index()

# %%
df_spolis_23_sum = df_spolis_23_sum.loc[df_spolis_23_sum['SBASISLOON']>0]


# %%
# Create hourly wage column by dividing total wage by total hours. 
df_spolis_23_sum['UURLOON'] = df_spolis_23_sum['SBASISLOON'] / df_spolis_23_sum['SREGULIEREUREN']
df_spolis_23_sum['UURLOON_BB'] = df_spolis_23_sum['LOON_BB'] / df_spolis_23_sum['SREGULIEREUREN']

# %%
# Merge spolis and gba and select relevant columns

df_gbaspolis_23 = pd.merge(df_gba_23, df_spolis_23_sum, on='RINPERSOON', how='inner')
df_gbaspolis_23 = df_gbaspolis_23[['leeftijd','SBASISLOON','UURLOON','LOON_BB','LOON_OW','LOON_ALL', 'UURLOON_BB']]

df_gbaspolis_23 = df_gbaspolis_23.loc[df_gbaspolis_23['leeftijd']>18]


# %%
# Create dataframes for different age groups.

df_gbaspolis_23_62 =  df_gbaspolis_23.loc[(df_gbaspolis_23['leeftijd']==62)]
df_gbaspolis_23_62.dropna(inplace=True)

df_gbaspolis_23_63 =  df_gbaspolis_23.loc[(df_gbaspolis_23['leeftijd']==63)]
df_gbaspolis_23_63.dropna(inplace=True)

df_gbaspolis_23_64 =  df_gbaspolis_23.loc[(df_gbaspolis_23['leeftijd']==64)]
df_gbaspolis_23_64.dropna(inplace=True)





# %%
get_column_statistics(df_gbaspolis_23, ['SBASISLOON', 'UURLOON','LOON_BB', 'LOON_OW', 'LOON_ALL', 'UURLOON_BB'])

# %%
with pd.ExcelWriter('rvu_analyse_update.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    get_column_statistics(df_gbaspolis_23_64, ['SBASISLOON', 'UURLOON','LOON_BB', 'LOON_OW', 'LOON_ALL', 'UURLOON_BB']).to_excel(writer,sheet_name='inkomen 64 jaar', startrow=1, index=False, header=True)

# %%
get_column_statistics(df_gbaspolis_23_64, ['SBASISLOON', 'UURLOON','LOON_BB', 'LOON_OW', 'LOON_ALL', 'UURLOON_BB'])

# %%
# Create graph for wage distribution for people 63 years of age. 

graph_data = df_gbaspolis_23_63.loc[(df_gbaspolis_23_63['LOON_ALL']<180000) & (df_gbaspolis_23_63['LOON_ALL']>0)]['LOON_ALL']
plt.figure(figsize=(12,6))
plt.hist(graph_data, color = "lightblue", bins=200, edgecolor = 'black', density=True)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_loon_63 = df_gbaspolis_23_63['LOON_ALL'].mean()
median_loon_63 = df_gbaspolis_23_63['LOON_ALL'].median()
mode_loon_63 = statistics.mode(df_gbaspolis_23_63['LOON_ALL'])
CPB_mode_loon_63 = (mean_loon_63/100)*79

plt.axvline(mean_loon_63, color='blue', linestyle = '--', label=f'mean: {mean_loon_63:.2f}')
plt.axvline(median_loon_63, color='purple', linestyle = '--', label=f'median: {median_loon_63:.2f}')
plt.axvline(mode_loon_63, color='green', linestyle = '--', label=f'mode: {mode_loon_63:.2f}')
#plt.axvline(CPB_mode_loon_63, color='orange', linestyle = '--', label=f'CPB mode: {CPB_mode_loon_63:.2f}')

plt.xlabel("Loon")
plt.ylabel("Density")

plt.title("verdeling loon (inclusief bb) 63-jarigen")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

g63w = plt.gcf()


# %%
# Create graph for distribution of hourly wages for people 63 years of age. 

graph_data = df_gbaspolis_23_63.loc[df_gbaspolis_23_63['UURLOON_BB'] <120]['UURLOON_BB']
plt.figure(figsize=(12,6))
plt.hist(graph_data, color = "lightblue", bins=200, edgecolor = 'black', density=True)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_uurloon_63 = df_gbaspolis_23_63['UURLOON_BB'].mean()
median_uurloon_63 = df_gbaspolis_23_63['UURLOON_BB'].median()
mode_uurloon_63 = statistics.mode(df_gbaspolis_23_63['UURLOON_BB'])
CPB_modeuurloon_63 = (mean_uurloon_63/100)*79

plt.axvline(mean_uurloon_63, color='blue', linestyle = '--', label=f'mean: {mean_uurloon_63:.2f}')
plt.axvline(median_uurloon_63, color='purple', linestyle = '--', label=f'median: {median_uurloon_63:.2f}')
plt.axvline(mode_uurloon_63, color='green', linestyle = '--', label=f'mode: {mode_uurloon_63:.2f}')
#plt.axvline(CPB_modeuurloon_63, color='orange', linestyle = '--', label=f'CPB mode: {CPB_modeuurloon_63:.2f}')

plt.xlabel('Uurloon')
plt.ylabel("Density")

plt.title("verdeling uurloon 63-jarigen")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

g63hw = plt.gcf()

# %%
# Create graph for wage distribution for total. 

graph_data = df_gbaspolis_23.loc[(df_gbaspolis_23['LOON_ALL']<180000) & (df_gbaspolis_23['LOON_ALL']>0)]['LOON_ALL']
plt.figure(figsize=(12,6))
plt.hist(graph_data, color = "lightblue", bins=200, edgecolor = 'black', density=True)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_loon = df_gbaspolis_23['LOON_ALL'].mean()
median_loon = df_gbaspolis_23['LOON_ALL'].median()
mode_loon = statistics.mode(df_gbaspolis_23['LOON_ALL'])
CPB_mode = (mean_loon/100)*79

plt.axvline(mean_loon, color='blue', linestyle = '--', label=f'mean: {mean_loon:.2f}')
plt.axvline(median_loon, color='purple', linestyle = '--', label=f'median: {median_loon:.2f}')
plt.axvline(mode_loon, color='green', linestyle = '--', label=f'mode: {mode_loon:.2f}')
#plt.axvline(CPB_mode, color='orange', linestyle = '--', label=f'CPB mode: {CPB_mode:.2f}')

plt.axvline(mean_loon_63, color='black', linestyle = '--', label=f'mean 63: {mean_loon_63:.2f}')
plt.axvline(median_loon_63, color='orange', linestyle = '--', label=f'median 63: {median_loon_63:.2f}')

plt.xlabel("Loon")
plt.ylabel("Density")

plt.title("verdeling loon (inclusief bb) totaal")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

gtotw = plt.gcf()



# %%
# Create graph for distribution of hourly wages total. 

graph_data = df_gbaspolis_23.loc[df_gbaspolis_23['UURLOON_BB'] <120]['UURLOON_BB']
plt.figure(figsize=(12,6))
plt.hist(graph_data, color = "lightblue", bins=200, edgecolor = 'black', density=True)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_uurloon = df_gbaspolis_23['UURLOON_BB'].mean()
median_uurloon = df_gbaspolis_23['UURLOON_BB'].median()
mode_uurloon = statistics.mode(df_gbaspolis_23['UURLOON_BB'])
CPB_mode_uurloon = (mean_uurloon/100)*79

plt.axvline(mean_uurloon, color='blue', linestyle = '--', label=f'mean: {mean_uurloon:.2f}')
plt.axvline(median_uurloon, color='purple', linestyle = '--', label=f'median: {median_uurloon:.2f}')
plt.axvline(mode_uurloon, color='green', linestyle = '--', label=f'mode: {mode_uurloon:.2f}')
#plt.axvline(CPB_mode_uurloon, color='orange', linestyle = '--', label=f'CPB mode: {CPB_mode_uurloon:.2f}')

plt.axvline(mean_uurloon_63, color='black', linestyle = '--', label=f'mean 63: {mean_uurloon_63:.2f}')
plt.axvline(median_uurloon_63, color='orange', linestyle = '--', label=f'median 63: {median_uurloon_63:.2f}')

plt.xlabel('Uurloon')
plt.ylabel("Density")

plt.title("verdeling uurloon totaal")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

gtothw = plt.gcf()



# %%
# create income class variable, and groupby income class.

bins_uurloon = [-float('inf'),0,5,10,15,20,25,30,35,40,45,50,55,60,65,70,75,80,85,90,95,100,float('inf')]
labels_uurloon = ['negative','0-4.99', '5-9.99', '10-14.99', '15-19.99', '20-24.99', '25-29.99', '30-34.99', '35-39.99', '40-44.99',
                    '45-49.99', '50-54.99', '55-59.99', '60-64.99', '65-69.99', '70-74.99', '75-79.99', '80-84.99', 
                    '85-89.99', '90-94.99', '95-99.99', 'meer dan 100']
df_gbaspolis_23['INCOME_CLASS'] = pd.cut(df_gbaspolis_23['UURLOON'], bins=bins_uurloon, labels=labels_uurloon, right=False)

df_gbaspolis_23.groupby('INCOME_CLASS').size()

# %%


# %%
# Create graph for distribution of hourly wages for people 62 years of age. 

graph_data = df_gbaspolis_23_62.loc[df_gbaspolis_23_62['UURLOON'] <120]['UURLOON']
plt.figure(figsize=(12,6))
plt.hist(graph_data, color = "lightblue", bins=200, edgecolor = 'black', density=True)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_uurloon_62 = df_gbaspolis_23_62['UURLOON'].mean()
median_uurloon_62 = df_gbaspolis_23_62['UURLOON'].median()
mode_uurloon_62 = statistics.mode(df_gbaspolis_23_62['UURLOON'])
CPB_mode_uurloon_62 = (mean_uurloon_62/100)*79

plt.axvline(mean_uurloon_62, color='blue', linestyle = '--', label=f'mean: {mean_uurloon_62:.2f}')
plt.axvline(median_uurloon_62, color='purple', linestyle = '--', label=f'median: {median_uurloon_62:.2f}')
plt.axvline(mode_uurloon_62, color='green', linestyle = '--', label=f'mode: {mode_uurloon_62:.2f}')
#plt.axvline(CPB_mode_uurloon_62, color='orange', linestyle = '--', label=f'CPB mode: {CPB_mode_uurloon_62:.2f}')

plt.xlabel('Uurloon')
plt.ylabel("Density")

plt.title("verdeling uurloon 62-jarigen")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

g62hw = plt.gcf()



# %%
# Create graph for wage distribution for people 62 years of age. 

graph_data = df_gbaspolis_23_62.loc[(df_gbaspolis_23_62['LOON_ALL']<180000) & (df_gbaspolis_23_62['LOON_ALL']>0)]['LOON_ALL']
plt.figure(figsize=(12,6))
plt.hist(graph_data, color = "lightblue", bins=200, edgecolor = 'black', density=True)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_loon_62 = df_gbaspolis_23_62['LOON_ALL'].mean()
median_loon_62 = df_gbaspolis_23_62['LOON_ALL'].median()
mode_loon_62 = statistics.mode(df_gbaspolis_23_62['LOON_ALL'])
CPB_mode_loon_62 = (mean_loon_62/100)*79

plt.axvline(mean_loon_62, color='blue', linestyle = '--', label=f'mean: {mean_loon_62:.2f}')
plt.axvline(median_loon_62, color='purple', linestyle = '--', label=f'median: {median_loon_62:.2f}')
plt.axvline(mode_loon_62, color='green', linestyle = '--', label=f'mode: {mode_loon_62:.2f}')
#plt.axvline(CPB_mode_loon_62, color='orange', linestyle = '--', label=f'CPB mode: {CPB_mode_loon_62:.2f}')

plt.xlabel("Loon")
plt.ylabel("Density")

plt.title("verdeling loon (inclusief bb) 62-jarigen")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

g62w = plt.gcf()


# %%
# Create graph for distribution of hourly wages for people 64 years of age. 

graph_data = df_gbaspolis_23_64.loc[df_gbaspolis_23_64['UURLOON_BB'] <120]['UURLOON_BB']
plt.figure(figsize=(12,6))
counts, bin_edges, _ = plt.hist(graph_data, color = "lightblue", bins=100, edgecolor = 'black', density=False)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

#sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_uurloon_64 = df_gbaspolis_23_64['UURLOON_BB'].mean()
median_uurloon_64 = df_gbaspolis_23_64['UURLOON_BB'].median()
mode_loon_64 = statistics.mode(df_gbaspolis_23_64['UURLOON_BB'])

plt.axvline(mean_uurloon_64, color='blue', linestyle = '--', label=f'mean: {mean_uurloon_64:.2f}')
plt.axvline(median_uurloon_64, color='purple', linestyle = '--', label=f'median: {median_uurloon_64:.2f}')
plt.axvline(mode_loon_64, color='green', linestyle = '--', label=f'mode: {mode_loon_64:.2f}')
#plt.axvline(CPB_mode_uurloon_64, color='orange', linestyle = '--', label=f'CPB mode: {CPB_mode_uurloon_64:.2f}')

plt.xlabel('UURLOON_BB')
plt.ylabel("Aantal")

plt.title("Graph 64_2: verdeling uurloon (UURLOON_BB, max 120 euro) 64-jarigen")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

plt.savefig('g_loon_all_64.png', dpi=400)



# %%
g_data=pd.DataFrame({
    'bin_start': bin_edges[:-1],
    'bin_end':bin_edges[1:],
    'count': counts
})

with pd.ExcelWriter('rvu_analyse_update.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    g_data.to_excel(writer,sheet_name='inkomen 64 jaar', startrow=12, index=False, header=True)
    

# %%
# Create graph for wage distribution for people 64 years of age. 

graph_data = df_gbaspolis_23_64.loc[(df_gbaspolis_23_64['LOON_ALL']<150000) & (df_gbaspolis_23_64['LOON_ALL']>0)]['LOON_ALL']
plt.figure(figsize=(12,6))
counts, bin_edges, _ = plt.hist(graph_data, color = "lightblue", bins=120, edgecolor = 'black', density=False)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

#sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_loon_64 = df_gbaspolis_23_64['LOON_ALL'].mean()
median_loon_64 = df_gbaspolis_23_64['LOON_ALL'].median()
mode_loon_64 = statistics.mode(df_gbaspolis_23_64['LOON_ALL'])


plt.axvline(mean_loon_64, color='blue', linestyle = '--', label=f'mean: {mean_loon_64:.2f}')
plt.axvline(median_loon_64, color='purple', linestyle = '--', label=f'median: {median_loon_64:.2f}')
plt.axvline(mode_loon_64, color='green', linestyle = '--', label=f'mode: {mode_loon_64:.2f}')
#plt.axvline(CPB_mode_loon_64, color='orange', linestyle = '--', label=f'CPB mode: {CPB_mode_loon_64:.2f}')

plt.xlabel("LOON_ALL")
plt.ylabel("Aantal")

plt.title("Graph 64_1: verdeling loon (LOON_ALL, max 150k) 64-jarigen")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

plt.savefig('g_loon_all_64b.png', dpi=400)



# %%
g_data=pd.DataFrame({
    'bin_start': bin_edges[:-1],
    'bin_end':bin_edges[1:],
    'count': counts
})

with pd.ExcelWriter('rvu_analyse_update.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    g_data.to_excel(writer,sheet_name='inkomen 64 jaar', startrow=12, index=False, header=True)
    

# %%
df_gba_23['leeftijd'].max()

# %%

# Writing basic statistics and graphs to excel file

excel_path = 'output.xlsx'
writer = pd.ExcelWriter(excel_path, engine='openpyxl')

(get_column_statistics(df_gbaspolis_23, ['SBASISLOON','LOON_BB', 'LOON_OW', 'LOON_ALL', 'UURLOON', 'UURLOON_BB'])).to_excel(writer, sheet_name='totaal')

(get_column_statistics(df_gbaspolis_23_62, ['SBASISLOON','LOON_BB', 'LOON_OW', 'LOON_ALL', 'UURLOON', 'UURLOON_BB'])).to_excel(writer, sheet_name='62 jaar')

(get_column_statistics(df_gbaspolis_23_63, ['SBASISLOON','LOON_BB', 'LOON_OW', 'LOON_ALL', 'UURLOON', 'UURLOON_BB'])).to_excel(writer, sheet_name='63 jaar')

(get_column_statistics(df_gbaspolis_23_64, ['SBASISLOON','LOON_BB', 'LOON_OW', 'LOON_ALL', 'UURLOON', 'UURLOON_BB'])).to_excel(writer, sheet_name='64 jaar')

writer.close()
wb = load_workbook(excel_path)

img62w = io.BytesIO()
g62w.savefig(img62w, format='png')
img62w.seek(0)

img62hw = io.BytesIO()
g62hw.savefig(img62hw, format='png')
img62hw.seek(0)

img63w = io.BytesIO()
g63w.savefig(img63w, format='png')
img63w.seek(0)

img63hw = io.BytesIO()
g63hw.savefig(img63hw, format='png')
img63hw.seek(0)

img64w = io.BytesIO()
g64w.savefig(img64w, format='png')
img64w.seek(0)

img64hw = io.BytesIO()
g64hw.savefig(img64hw, format='png')
img64hw.seek(0)

imgtotw = io.BytesIO()
gtotw.savefig(imgtotw, format='png')
imgtotw.seek(0)

imgtothw = io.BytesIO()
gtothw.savefig(imgtothw, format='png')
imgtothw.seek(0)

wb['totaal'].add_image(Image(imgtotw),'D9')
wb['totaal'].add_image(Image(imgtothw),'X9')

wb['62 jaar'].add_image(Image(img62w),'D9')
wb['62 jaar'].add_image(Image(img62hw),'X9')

wb['63 jaar'].add_image(Image(img63w),'D9')
wb['63 jaar'].add_image(Image(img63hw),'X9')

wb['64 jaar'].add_image(Image(img64w),'D9')
wb['64 jaar'].add_image(Image(img64hw),'X9')

wb.save(excel_path)


# %%
#Load INPATAB 2022
fpath_inpatab_22 = "G:\InkomenBestedingen\INPATAB\INPA2022TABV2.sav" # set file path inpatab 2022
metadata_inpatab_22 = prt.read_sav(fpath_inpatab_22, metadataonly=True) # load metadata inpatab 2022
columns_to_load_inpatab_22 = ['RINPERSOON','INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB','INPPOSHHK'] #select columns to load from inpatab 2022
df_inpatab_22 = pd.read_spss (fpath_inpatab_22, usecols=columns_to_load_inpatab_22) #load inpatab 2022 data

#Convert income columns to numeric (float)
df_inpatab_22[['INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB']] = df_inpatab_22[['INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB']].apply(pd.to_numeric, errors='coerce')

# create UITWERK column
df_inpatab_22['UITWERK'] = df_inpatab_22['INPT1000WER'] + df_inpatab_22['INPT1020AMB']


#load SPOLISBUS 2022
fpath_spolis_22 = r"G:\Spolis\SPOLISBUS\2022\SPOLISBUS2022V5.sav" # set file path spolisbus 2022
metadata_spolis_22 = prt.read_sav(fpath_spolis_22, metadataonly=True) # load metadata spolisbus 2022
columns_to_load_spolisbus = ['RINPERSOON', 'SREGULIEREUREN'] #select columns to load from spolisbus 2022
df_spolis_22= pd.read_spss(fpath_spolis_22, usecols=columns_to_load_spolisbus) #load spolisbus data 2022


#Load GBAPERSOONKTAB 2022

fpath_gba_22 = "G:\Bevolking\GBAPERSOONKTAB\GBAPERSOONKTAB2022V1.sav" # set file path GBAPERSOONKTAB 2023
df_gba_22 = pd.read_spss(fpath_gba_22) # load GBAPERSOONKTAB data 

# Create age column GBA 2022
df_gba_22['GBAGEBOORTEJAAR'] = pd.to_numeric(df_gba_23['GBAGEBOORTEJAAR'], errors='coerce') # convert to numeric
df_gba_22['leeftijd'] = 2022 - df_gba_22['GBAGEBOORTEJAAR'] # calculate age and create column




# %%

#load VEHTAB 2022
fpath_veh_22 = r"G:\InkomenBestedingen\VEHTAB\VEH2022TABV1.sav" # set file path VEHTAB 2022
metadata_veh_22 = prt.read_sav(fpath_veh_22, metadataonly=True) # load metadata VEHTAB 2022
columns_to_load_veh_22 = ['RINPERSOONHKW', 'VEHW1000VERH', 'VEHW1121WONH','VEHW1110FINH','VEHW1111BANH'] #select columns to load from VEHTAB 2022
df_veh_22 = pd.read_spss(fpath_veh_22, usecols=columns_to_load_veh_22) #load VEHTAB 2022 data 


#load koppelbestand huishoudens -> personen 2022
fpath_koppel_veh_2022 = r"G:\InkomenBestedingen\VEHTAB\KOPPELPERSOONHUISHOUDEN2022V1.sav"
metadata_koppel_veh_22 = prt.read_sav(fpath_koppel_veh_2022, metadataonly=True)
columns_to_load__koppel_veh_22 = ['RINPERSOON', 'RINPERSOONHKW']
df_veh_koppel_22 = pd.read_spss(fpath_koppel_veh_2022, usecols=columns_to_load__koppel_veh_22)

# %%
df1 = pd.merge(df_gba_22, df_inpatab_22, on='RINPERSOON', how='inner')

# %%
df1_64 = df1.loc[df1['leeftijd']==64]

# %%
# Create graph for wage distribution for people 64 years of age. 

graph_data = df1_64.loc[(df1_64['UITWERK']<150000) & (df1_64['UITWERK']>1)]['UITWERK']
plt.figure(figsize=(12,6))
counts, bin_edges, _ = plt.hist(graph_data, color = "lightblue", bins=120, edgecolor = 'black', density=False)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

#sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_loon_64 = df1_64.loc[(df1_64['UITWERK']>1)]['UITWERK'].mean()
median_loon_64 = df1_64.loc[(df1_64['UITWERK']>1)]['UITWERK'].median()
mode_loon_64 = statistics.mode((df1_64.loc[df1_64['UITWERK']>1])['UITWERK'])


plt.axvline(mean_loon_64, color='blue', linestyle = '--', label=f'mean: {mean_loon_64:.2f}')
plt.axvline(median_loon_64, color='purple', linestyle = '--', label=f'median: {median_loon_64:.2f}')
plt.axvline(mode_loon_64, color='green', linestyle = '--', label=f'mode: {mode_loon_64:.2f}')
#plt.axvline(CPB_mode_loon_64, color='orange', linestyle = '--', label=f'CPB mode: {CPB_mode_loon_64:.2f}')

plt.xlabel("loon UITWERK")
plt.ylabel("Aantal")

plt.title("Graph 64_3: verdeling loon (UITWERK, max 150k) 64-jarigen")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

plt.savefig('g_UITWERK_64.png', dpi=400)



# %%
g_data=pd.DataFrame({
    'bin_start': bin_edges[:-1],
    'bin_end':bin_edges[1:],
    'count': counts
})

with pd.ExcelWriter('rvu_analyse_update.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    g_data.to_excel(writer,sheet_name='inkomen 64 jaar UITWERK', startrow=8, index=False, header=True)
    

# %%
with pd.ExcelWriter('rvu_analyse_update.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
   get_column_statistics(df1_64.loc[df1_64['UITWERK']>1], ['UITWERK']).to_excel(writer,sheet_name='inkomen 64 jaar UITWERK', startrow=1, index=False, header=True)

# %%
get_column_statistics(df1_64, ['UITWERK'])

# %%
df_v = pd.merge(df_veh_22, df_veh_koppel_22, on='RINPERSOONHKW', how = 'inner')

df2 = pd.merge(df_gba_22, df_v, on='RINPERSOON', how='inner')

# %%
df2_64 = df2.loc[df2['leeftijd']==64]

# %%
df2_64['VEHW1000VERH'] = pd.to_numeric(df2_64['VEHW1000VERH'], errors='coerce').astype(float)
df2_64['VEHW1121WONH'] = pd.to_numeric(df2_64['VEHW1121WONH'], errors='coerce').astype(float)
df2_64['VEHW1111BANH'] = pd.to_numeric(df2_64['VEHW1111BANH'], errors='coerce').astype(float)
df2_64['VEHW1110FINH'] = pd.to_numeric(df2_64['VEHW1110FINH'], errors='coerce').astype(float)

# %%
df2_64 = df2_64.dropna()

# %%
with pd.ExcelWriter('rvu_analyse_update.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    get_column_statistics(df2_64,['VEHW1000VERH','VEHW1121WONH','VEHW1111BANH','VEHW1110FINH']).to_excel(writer,sheet_name='vermogen 64 jaar', startrow=1, index=False, header=True)



# %%
# Create graph for wage distribution for people 64 years of age. 

graph_data = df2_64.loc[(df2_64['VEHW1000VERH']>-150000) & (df2_64['VEHW1000VERH']<1000000)]['VEHW1000VERH']
plt.figure(figsize=(16,6))
counts, bin_edges, _ = plt.hist(graph_data, color = "lightblue", bins=120, edgecolor = 'black', density=False)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

#sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_verm_64 = df2_64['VEHW1000VERH'].mean()
median_verm_64 = df2_64['VEHW1000VERH'].median()
mode_verm_64 = statistics.mode(df2_64['VEHW1000VERH'])


plt.axvline(mean_verm_64, color='blue', linestyle = '--', label=f'mean: {mean_verm_64:.2f}')
plt.axvline(median_verm_64, color='purple', linestyle = '--', label=f'median: {median_verm_64:.2f}')
plt.axvline(mode_verm_64, color='green', linestyle = '--', label=f'mode: {mode_verm_64:.2f}')
#plt.axvline(CPB_mode_loon_64, color='orange', linestyle = '--', label=f'CPB mode: {CPB_mode_loon_64:.2f}')

plt.xlabel("totaal vermogen")
plt.ylabel("Aantal")

plt.title("Graph 64_4: verdeling vrmogen (totale vermogen, min -100k, max 1.m) 64-jarigen")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

plt.savefig('g_verm_64.png', dpi=400)



# %%
g_data=pd.DataFrame({
    'bin_start': bin_edges[:-1],
    'bin_end':bin_edges[1:],
    'count': counts
})

with pd.ExcelWriter('rvu_analyse_update.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    g_data.to_excel(writer,sheet_name='vermogen 64 jaar', startrow=10, index=False, header=True)
    

# %%
# Create graph for wage distribution for people 64 years of age. 

graph_data = df2_64.loc[(df2_64['VEHW1111BANH']>-150000) & (df2_64['VEHW1111BANH']<300000)]['VEHW1111BANH']
plt.figure(figsize=(16,6))
counts, bin_edges, _ = plt.hist(graph_data, color = "lightblue", bins=120, edgecolor = 'black', density=False)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

#sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_verm_64 = df2_64['VEHW1111BANH'].mean()
median_verm_64 = df2_64['VEHW1111BANH'].median()
mode_verm_64 = statistics.mode(df2_64['VEHW1111BANH'])


plt.axvline(mean_verm_64, color='blue', linestyle = '--', label=f'mean: {mean_verm_64:.2f}')
plt.axvline(median_verm_64, color='purple', linestyle = '--', label=f'median: {median_verm_64:.2f}')
plt.axvline(mode_verm_64, color='green', linestyle = '--', label=f'mode: {mode_verm_64:.2f}')
#plt.axvline(CPB_mode_loon_64, color='orange', linestyle = '--', label=f'CPB mode: {CPB_mode_loon_64:.2f}')

plt.xlabel("bank- en spaartegoeden")
plt.ylabel("Aantal")

plt.title("Graph 64_5: verdeling vrmogen (bank- en spaartegoeden, min -100k, max 1.m) 64-jarigen")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()

plt.savefig('g_verm_64b.png', dpi=400)



# %%
g_data=pd.DataFrame({
    'bin_start': bin_edges[:-1],
    'bin_end':bin_edges[1:],
    'count': counts
})

# %%
g_data=pd.DataFrame({
    'bin_start': bin_edges[:-1],
    'bin_end':bin_edges[1:],
    'count': counts
})

with pd.ExcelWriter('rvu_analyse_update.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    g_data.to_excel(writer,sheet_name='vermogen 64 jaar', startrow=10,startcol=18, index=False, header=True)
    

# %%



