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
#load RVU data

fpath_RVU = "L:\9765_GLV13815_ResultatenCBKV1.sav"
df_RVU = pd.read_spss(fpath_RVU) #load RVU data

# %%
df_RVU

# %%
#Load GBAPERSOONKTAB 2023

fpath_gba_23 = "G:\Bevolking\GBAPERSOONKTAB\GBAPERSOONKTAB2023V1.sav" # set file path GBAPERSOONKTAB 2023
df_gba_23 = pd.read_spss(fpath_gba_23) # load GBAPERSOONKTAB data 

# Create age column GBA 2023
df_gba_23['GBAGEBOORTEJAAR'] = pd.to_numeric(df_gba_23['GBAGEBOORTEJAAR'], errors='coerce') # convert to numeric
df_gba_23['leeftijd'] = 2024 - df_gba_23['GBAGEBOORTEJAAR'] # calculate age and create column

df_gba_23 = df_gba_23[['RINPERSOON','GBAGESLACHT','GBAGEBOORTEJAAR','leeftijd']] #retain only needed columns

# %%
#LOAD SPOLIBUS DATA

#load SPOLISBUS 2020
fpath_spolis_20 = r"G:\Spolis\SPOLISBUS\2020\SPOLISBUS2020V5.sav" # set file path spolisbus 2020
metadata_spolis_20 = prt.read_sav(fpath_spolis_20, metadataonly=True) # load metadata spolisbus 2020
columns_to_load_spolisbus = ['RINPERSOON', 'SBASISLOON','SLNOWRK','SBASISUREN', 'SREGULIEREUREN', 'SBIJZONDEREBELONING'] #select columns to load from spolisbus 2020
df_spolis_20 = pd.read_spss(fpath_spolis_20, usecols=columns_to_load_spolisbus) #load spolisbus data 2020

df_spolis_20['LOON_BB'] = df_spolis_20['SBASISLOON'] + df_spolis_20['SBIJZONDEREBELONING']
df_spolis_20['LOON_OW'] = df_spolis_20['SBASISLOON'] + df_spolis_20['SLNOWRK']
df_spolis_20['LOON_ALL'] = df_spolis_20['SBASISLOON'] + df_spolis_20['SBIJZONDEREBELONING'] + df_spolis_20['SLNOWRK']

#load SPOLISBUS 2021
fpath_spolis_21 = r"G:\Spolis\SPOLISBUS\2021\SPOLISBUS2021V5.sav"# set file path spolisbus 2021
metadata_spolis_21 = prt.read_sav(fpath_spolis_21, metadataonly=True) # load metadata spolisbus 2021
columns_to_load_spolisbus = ['RINPERSOON', 'SBASISLOON','SLNOWRK','SBASISUREN', 'SREGULIEREUREN', 'SBIJZONDEREBELONING'] #select columns to load from spolisbus 2021
df_spolis_21 = pd.read_spss(fpath_spolis_21, usecols=columns_to_load_spolisbus) #load spolisbus data 2021

df_spolis_21['LOON_BB'] = df_spolis_21['SBASISLOON'] + df_spolis_21['SBIJZONDEREBELONING']
df_spolis_21['LOON_OW'] = df_spolis_21['SBASISLOON'] + df_spolis_21['SLNOWRK']
df_spolis_21['LOON_ALL'] = df_spolis_21['SBASISLOON'] + df_spolis_21['SBIJZONDEREBELONING'] + df_spolis_21['SLNOWRK']

#load SPOLISBUS 2022
fpath_spolis_22 = r"G:\Spolis\SPOLISBUS\2022\SPOLISBUS2022V5.sav" # set file path spolisbus 2022
metadata_spolis_22 = prt.read_sav(fpath_spolis_22, metadataonly=True) # load metadata spolisbus 2022
columns_to_load_spolisbus = ['RINPERSOON', 'SBASISLOON','SLNOWRK','SBASISUREN', 'SREGULIEREUREN', 'SBIJZONDEREBELONING'] #select columns to load from spolisbus 2022
df_spolis_22= pd.read_spss(fpath_spolis_22, usecols=columns_to_load_spolisbus) #load spolisbus data 2022

df_spolis_22['LOON_BB'] = df_spolis_22['SBASISLOON'] + df_spolis_22['SBIJZONDEREBELONING']
df_spolis_22['LOON_OW'] = df_spolis_22['SBASISLOON'] + df_spolis_22['SLNOWRK']
df_spolis_22['LOON_ALL'] = df_spolis_22['SBASISLOON'] + df_spolis_22['SBIJZONDEREBELONING'] + df_spolis_22['SLNOWRK']

#load SPOLISBUS 2023
fpath_spolis = r"G:\Spolis\SPOLISBUS\2023\SPOLISBUS2023V4.sav" # set file path spolisbus 2023
metadata_spolis = prt.read_sav(fpath_spolis, metadataonly=True) # load metadata spolisbus 2023
columns_to_load_spolisbus = ['RINPERSOON', 'SBASISLOON','SLNOWRK','SBASISUREN', 'SREGULIEREUREN', 'SBIJZONDEREBELONING'] #select columns to load from spolisbus 2023
df_spolis_23 = pd.read_spss(fpath_spolis, usecols=columns_to_load_spolisbus) #load spolisbus data 2023

df_spolis_23['LOON_BB'] = df_spolis_23['SBASISLOON'] + df_spolis_23['SBIJZONDEREBELONING']
df_spolis_23['LOON_OW'] = df_spolis_23['SBASISLOON'] + df_spolis_23['SLNOWRK']
df_spolis_23['LOON_ALL'] = df_spolis_23['SBASISLOON'] + df_spolis_23['SBIJZONDEREBELONING'] + df_spolis_23['SLNOWRK']





# %%
#LOAD INPATAB DATA

#Load INPATAB 2022
fpath_inpatab_22 = "G:\InkomenBestedingen\INPATAB\INPA2022TABV2.sav" # set file path inpatab 2022
metadata_inpatab_22 = prt.read_sav(fpath_inpatab_22, metadataonly=True) # load metadata inpatab 2022
columns_to_load_inpatab_22 = ['RINPERSOON','INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB','INPPOSHHK'] #select columns to load from inpatab 2022
df_inpatab_22 = pd.read_spss (fpath_inpatab_22, usecols=columns_to_load_inpatab_22) #load inpatab 2022 data

#Convert income columns to numeric (float)
df_inpatab_22[['INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB']] = df_inpatab_22[['INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB']].apply(pd.to_numeric, errors='coerce')

# create UITWERK column
df_inpatab_22['UITWERK'] = df_inpatab_22['INPT1000WER'] + df_inpatab_22['INPT1020AMB']

#Load INPATAB 2021
fpath_inpatab_21 = "G:\InkomenBestedingen\INPATAB\INPA2021TABV3.sav" # set file path inpatab 2021
metadata_inpatab_21 = prt.read_sav(fpath_inpatab_21, metadataonly=True) # load metadata inpatab 2021
columns_to_load_inpatab_21 = ['RINPERSOON','INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB','INPPOSHHK'] #select columns to load from inpatab 2021
df_inpatab_21 = pd.read_spss (fpath_inpatab_21, usecols=columns_to_load_inpatab_21) #load inpatab 2021 data

#Convert income columns to numeric (float)
df_inpatab_21[['INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB']] = df_inpatab_21[['INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB']].apply(pd.to_numeric, errors='coerce')

# create UITWERK column
df_inpatab_21['UITWERK'] = df_inpatab_21['INPT1000WER'] + df_inpatab_21['INPT1020AMB']

#Load INPATAB 2020
fpath_inpatab_20 = "G:\InkomenBestedingen\INPATAB\INPA2020TABV3.sav" # set file path inpatab 2020
metadata_inpatab_20 = prt.read_sav(fpath_inpatab_20, metadataonly=True) # load metadata inpatab 2020
columns_to_load_inpatab_20 = ['RINPERSOON','INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB','INPPOSHHK'] #select columns to load from inpatab 2020
df_inpatab_20 = pd.read_spss (fpath_inpatab_20, usecols=columns_to_load_inpatab_20) #load inpatab 2020 data

#Convert income columns to numeric (float)
df_inpatab_20[['INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB']] = df_inpatab_20[['INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB']].apply(pd.to_numeric, errors='coerce')

# create UITWERK column
df_inpatab_20['UITWERK'] = df_inpatab_20['INPT1000WER'] + df_inpatab_20['INPT1020AMB']

# INPPERSBRUT is het bruto persoonlijk  inkomen (uit werk/eigen onderneming/ uitkeringe/toeslagen/etc.)
# INPPERSINK is het persoonlijk  inkomen (uit werk/eigen onderneming/ uitkeringe/toeslagen/etc.) met premies inkomensverzekering in mindering gebracht.
# INPPERSPRIM is het persoonlijk primair inkomen (inkomen uit werk en eigen inderneming)
# INPT1000WER is het loon van een werknemer/ inkomen uit arbeid
# INPT1020AMB is het loon van een ambtenaar/ inkomen uit arbeid ambtenaar

# %%
#load HOOGSTEOPLTAB 2022
fpath_opl_22 = r"G:\Onderwijs\HOOGSTEOPLTAB\2022\HOOGSTEOPL2022TABV1.sav" # set file path HOOGSTEOPLTAB 2022
metadata_opl = prt.read_sav(fpath_opl_22, metadataonly=True) # load metadata HOOGSTEOPLTAB 2022
columns_to_load_opl_22 = ['RINPERSOON', 'OPLNIVSOI2021AGG4HBmetNIRWO'] #select columns to load from HOOGSTEOPLTAB 2022
df_opl_22 = pd.read_spss(fpath_opl_22, usecols=columns_to_load_opl_22) #load HOOGSTEOPLTAB 2022


# %%
#LOAD VEHTAB DATA

#load VEHTAB 2022
fpath_veh_22 = r"G:\InkomenBestedingen\VEHTAB\VEH2022TABV1.sav" # set file path VEHTAB 2022
metadata_veh_22 = prt.read_sav(fpath_veh_22, metadataonly=True) # load metadata VEHTAB 2022
columns_to_load_veh_22 = ['RINPERSOONHKW', 'VEHW1000VERH', 'VEHW1121WONH','VEHW1110FINH','VEHW1111BANH'] #select columns to load from VEHTAB 2022
df_veh_22 = pd.read_spss(fpath_veh_22, usecols=columns_to_load_veh_22) #load VEHTAB 2022 data 

#load VEHTAB 2021
fpath_veh_21 = r"G:\InkomenBestedingen\VEHTAB\VEH2021TABV3.sav" # set file path VEHTAB 2021
metadata_veh_21 = prt.read_sav(fpath_veh_21, metadataonly=True) # load metadata VEHTAB 2021
columns_to_load_veh_21 = ['RINPERSOONHKW', 'VEHW1000VERH', 'VEHW1121WONH','VEHW1110FINH','VEHW1111BANH'] #select columns to load from VEHTAB 2021
df_veh_21 = pd.read_spss(fpath_veh_21, usecols=columns_to_load_veh_21) #load VEHTAB 2021 data 

#load VEHTAB 2020
fpath_veh_20 = r"G:\InkomenBestedingen\VEHTAB\VEH2020TABV3.sav" # set file path VEHTAB 2022
metadata_veh_20 = prt.read_sav(fpath_veh_20, metadataonly=True) # load metadata VEHTAB 2022
columns_to_load_veh_20 = ['RINPERSOONHKW', 'VEHW1000VERH', 'VEHW1121WONH','VEHW1110FINH','VEHW1111BANH'] #select columns to load from VEHTAB 2022
df_veh_20 = pd.read_spss(fpath_veh_20, usecols=columns_to_load_veh_20) #load VEHTAB 2022 data 

#load koppelbestand huishoudens -> personen 2022
fpath_koppel_veh_2022 = r"G:\InkomenBestedingen\VEHTAB\KOPPELPERSOONHUISHOUDEN2022V1.sav"
metadata_koppel_veh_22 = prt.read_sav(fpath_koppel_veh_2022, metadataonly=True)
columns_to_load__koppel_veh_22 = ['RINPERSOON', 'RINPERSOONHKW']
df_veh_koppel_22 = pd.read_spss(fpath_koppel_veh_2022, usecols=columns_to_load__koppel_veh_22)

#load koppelbestand huishoudens -> personen 2021
fpath_koppel_veh_2021 = r"G:\InkomenBestedingen\VEHTAB\KOPPELPERSOONHUISHOUDEN2021V2.sav"
metadata_koppel_veh_21 = prt.read_sav(fpath_koppel_veh_2021, metadataonly=True)
columns_to_load__koppel_veh_21 = ['RINPERSOON', 'RINPERSOONHKW']
df_veh_koppel_21 = pd.read_spss(fpath_koppel_veh_2021, usecols=columns_to_load__koppel_veh_21)

#load koppelbestand huishoudens -> personen 2020
fpath_koppel_veh_2020 = r"G:\InkomenBestedingen\VEHTAB\KOPPELPERSOONHUISHOUDEN2020V2.sav"
metadata_koppel_veh_20 = prt.read_sav(fpath_koppel_veh_2020, metadataonly=True)
columns_to_load__koppel_veh_20 = ['RINPERSOON', 'RINPERSOONHKW']
df_veh_koppel_20 = pd.read_spss(fpath_koppel_veh_2020, usecols=columns_to_load__koppel_veh_20)



# %%
#LOAD RVU INCOME
df_rvu_ink = pd.read_csv("F:\Desktop\Python project\Inkomen\Datafiles\RVU_INKOMEN.csv")

# %%
# For each unique ID number all registered wages are summed for each year. 

df_spolis_23_sum = df_spolis_23.groupby('RINPERSOON')[['SBASISLOON','SBASISUREN','LOON_BB', 'LOON_OW', 'LOON_ALL','SREGULIEREUREN']].sum().reset_index()
df_spolis_22_sum = df_spolis_22.groupby('RINPERSOON')[['SBASISLOON','SBASISUREN','LOON_BB', 'LOON_OW', 'LOON_ALL','SREGULIEREUREN']].sum().reset_index()
df_spolis_21_sum = df_spolis_21.groupby('RINPERSOON')[['SBASISLOON','SBASISUREN','LOON_BB', 'LOON_OW', 'LOON_ALL','SREGULIEREUREN']].sum().reset_index()
df_spolis_20_sum = df_spolis_20.groupby('RINPERSOON')[['SBASISLOON','SBASISUREN','LOON_BB', 'LOON_OW', 'LOON_ALL','SREGULIEREUREN']].sum().reset_index()

# %%
#Add year columns for wage data
df_spolis_23_sum['JAAR_LOON'] = 2023
df_spolis_22_sum['JAAR_LOON'] = 2022
df_spolis_21_sum['JAAR_LOON'] = 2021
df_spolis_20_sum['JAAR_LOON'] = 2020

#Add year columns for income data
df_inpatab_22['JAAR_INKOMEN'] = 2022
df_inpatab_21['JAAR_INKOMEN'] = 2021
df_inpatab_20['JAAR_INKOMEN'] = 2020

#Add year columns for wealth data
df_veh_22['JAAR_VERMOGEN'] = 2022
df_veh_21['JAAR_VERMOGEN'] = 2021
df_veh_20['JAAR_VERMOGEN'] = 2020

# %%
# Spolis dataframes are concatenated into a new dataframe df spolis complete (2020,2021,2022)
df_spolis_c = pd.concat([df_spolis_20_sum, df_spolis_21_sum, df_spolis_22_sum])

# %%
#Connect wealth data tot person ID's
df_veh_22_p = pd.merge(df_veh_22,df_veh_koppel_22, on='RINPERSOONHKW', how='inner')
df_veh_21_p = pd.merge(df_veh_21,df_veh_koppel_21, on='RINPERSOONHKW', how='inner')
df_veh_20_p = pd.merge(df_veh_20,df_veh_koppel_20, on='RINPERSOONHKW', how='inner')

# %%
# VEH dataframes are concatenated into a new dataframe df veh complete (2020,2021,2022)
df_veh_c = pd.concat([df_veh_20_p, df_veh_21_p, df_veh_22_p])

# %%
# INPATAB dataframes are concatenated into a new dataframe df ink complete (2020,2021,2022)
df_ink_c = pd.concat([df_inpatab_20, df_inpatab_21, df_inpatab_22])

# %%
#Merge RVU and GBA dataframes, keeping only RVU matches
df_RVU_gba = pd.merge(df_RVU,df_gba_23,how='left',on='RINPERSOON')

# %%
#Add spolis data
df_RVU_gba_spolis = pd.merge(df_RVU_gba, df_spolis_c,how='left', on='RINPERSOON')

# %%
# Keep rows with spolis income year before RVU
df_RVU_gba_spolis_f = df_RVU_gba_spolis.loc[df_RVU_gba_spolis['JAAR_LOON'] == df_RVU_gba_spolis['JAAR']-1]

# %%
#retain only people that didn't have RVU before 2021
df_RVU_gba_spolis_f_21 = df_RVU_gba_spolis_f.loc[df_RVU_gba_spolis_f['RVU']=='RVU pas in 2021']

# %%
#Convert RINPERSOON to numeric variable to enable merging
df_ink_c['RINPERSOON'] = df_ink_c['RINPERSOON'].astype(int)

# %%
df_RVU_gba_spolis_f_inkomen = pd.merge(df_RVU_gba_spolis_f_21,df_ink_c,how='left',on='RINPERSOON')

# %%
df_RVU_gba_spolis_f_inkomen_f = df_RVU_gba_spolis_f_inkomen.loc[df_RVU_gba_spolis_f_inkomen['JAAR_INKOMEN'] == df_RVU_gba_spolis_f_inkomen['JAAR']-1]

# %%
df_veh_c['RINPERSOON'] = df_veh_c['RINPERSOON'].astype(int) 

# %%
df_RVU_gba_spolis_f_inkomen_f_vermogen = pd.merge(df_RVU_gba_spolis_f_inkomen_f,df_veh_c,how='left',on='RINPERSOON')

# %%
df_RVU_gba_spolis_f_inkomen_f_vermogen_f = df_RVU_gba_spolis_f_inkomen_f_vermogen.loc[df_RVU_gba_spolis_f_inkomen_f_vermogen['JAAR_VERMOGEN'] == df_RVU_gba_spolis_f_inkomen_f_vermogen['JAAR']-1]

# %%
df_opl_22['RINPERSOON'] = df_opl_22['RINPERSOON'].astype(int)

# %%
df_RVU_gba_spolis_f_inkomen_f_vermogen_f_opl = pd.merge(df_RVU_gba_spolis_f_inkomen_f_vermogen_f,df_opl_22,how='left',on='RINPERSOON')

# %%
df_c = df_RVU_gba_spolis_f_inkomen_f_vermogen_f_opl

# %%
df_gba_23['RINPERSOON'] = df_gba_23['RINPERSOON'].astype(int)
df_c = pd.merge(df_c,df_gba_23, how='left', on='RINPERSOON')

# %%
df_c['VEHW1000VERH'] = pd.to_numeric(df_c['VEHW1000VERH'], errors='coerce').astype(float)
df_c['VEHW1121WONH'] = pd.to_numeric(df_c['VEHW1121WONH'], errors='coerce').astype(float)
df_c['VEHW1111BANH'] = pd.to_numeric(df_c['VEHW1111BANH'], errors='coerce').astype(float)
df_c['VEHW1110FINH'] = pd.to_numeric(df_c['VEHW1110FINH'], errors='coerce').astype(float)

# %%
df_c = df_c.loc[df_c['INPPERSINK']>0]

# %%
df_c['UURLOON'] = df_c['SBASISLOON']/df_c['SREGULIEREUREN']
df_c['UURLOON_BB'] = df_c['LOON_BB']/df_c["SREGULIEREUREN"]
df_c['UURLOON_UITWERK'] = df_c['UITWERK']/df_c['SREGULIEREUREN']

# %%
df_c['LEEFTIJD_RVU'] = df_c["JAAR"] - df_c["GBAGEBOORTEJAAR"]

# %%
df_c['EIGEN_WONING'] = np.where(df_c['VEHW1121WONH']==0,'nee','ja')

# %%
index_mapping = {2020:100,2021:102.1,2022:105.4,2023:111.8}

def adjust_wages(row, index_mapping):
    year_index = index_mapping[row['JAAR_LOON']]
    adjustment_factor = index_mapping[2023]/year_index
    return row['LOON_ALL']*adjustment_factor

df_c['LOON_ALL_IDX'] = df_c.apply(adjust_wages,axis=1, index_mapping=index_mapping)

# %%
def adjust_income(row, index_mapping):
    year_index = index_mapping[row['JAAR_INKOMEN']]
    adjustment_factor = index_mapping[2023]/year_index
    return row['UITWERK']*adjustment_factor

df_c['UITWERK_IDX'] = df_c.apply(adjust_income,axis=1, index_mapping=index_mapping)

# %%
def adjust_income(row, index_mapping):
    year_index = index_mapping[row['JAAR_LOON']]
    adjustment_factor = index_mapping[2023]/year_index
    return row['UURLOON']*adjustment_factor

df_c['UURLOON_IDX'] = df_c.apply(adjust_income,axis=1, index_mapping=index_mapping)

# %%
def adjust_income(row, index_mapping):
    year_index = index_mapping[row['JAAR_LOON']]
    adjustment_factor = index_mapping[2023]/year_index
    return row['UURLOON_BB']*adjustment_factor

df_c['UURLOON_BB_IDX'] = df_c.apply(adjust_income,axis=1, index_mapping=index_mapping)

# %%
def adjust_income(row, index_mapping):
    year_index = index_mapping[row['JAAR_LOON']]
    adjustment_factor = index_mapping[2023]/year_index
    return row['SBASISLOON']*adjustment_factor

df_c['SBASISLOON_IDX'] = df_c.apply(adjust_income,axis=1, index_mapping=index_mapping)

# %%
def adjust_income(row, index_mapping):
    year_index = index_mapping[row['JAAR_LOON']]
    adjustment_factor = index_mapping[2023]/year_index
    return row['LOON_BB']*adjustment_factor

df_c['LOON_BB_IDX'] = df_c.apply(adjust_income,axis=1, index_mapping=index_mapping)

# %%
def adjust_income(row, index_mapping):
    year_index = index_mapping[row['JAAR_LOON']]
    adjustment_factor = index_mapping[2023]/year_index
    return row['UURLOON_UITWERK']*adjustment_factor

df_c['UURLOON_UITWERK_IDX'] = df_c.apply(adjust_income,axis=1, index_mapping=index_mapping)

# %%
df_c.drop(['RINPERSOONS', 'RINPERSOONHKW','JAAR_VERMOGEN','JAAR_LOON', 'JAAR_INKOMEN','leeftijd_y','GBAGESLACHT_y'], axis=1, inplace=True)

# %%
df_c.info()

# %%
df_c.to_csv('rvu_compleet_def.csv', index=False)

# %%
get_column_statistics(df_c.loc[df_c['CAO']== 2948],["LOON_ALL", "SREGULIEREUREN","INPPERSINK","UITWERK","UURLOON","VEHW1000VERH","LEEFTIJD_RVU"])

# %%
df_c.loc[df_c['CAO']== 2948]['GBAGESLACHT_x'].value_counts()

# %%
df_c['LEEFTIJD_RVU'].value_counts()

# %%
df_c.groupby('CAO')['LOON_ALL'].count()

# %%
df_c.loc[df_c['LOON_ALL']>45782.500000].groupby('CAO')['RINPERSOON'].count()

# %%
(df_c.loc[df_c['LOON_ALL']>45782.500000]['CAO'].value_counts()/df_c['CAO'].value_counts()*100).sort_values(ascending=False).head(100)

# %%
graph_data = df_c.loc[df_c['CAO'] == 1636.0]['LOON_ALL']
plt.figure(figsize=(12,6))
plt.hist(graph_data, color = "lightblue", bins=100, edgecolor = 'black', density=True)

plt.grid(axis='x', color='lightgray', linestyle='--', linewidth =1, which='major')

sns.kdeplot(graph_data, color='red', linestyle = '-', linewidth=1)

mean_loon = df_c.loc[df_c['CAO'] == 1636.0]['LOON_ALL'].mean()
median_loon = df_c.loc[df_c['CAO'] == 1636.0]['LOON_ALL'].median()


plt.axvline(mean_loon, color='blue', linestyle = '--', label=f'mean: {mean_loon:.2f}')
plt.axvline(median_loon, color='purple', linestyle = '--', label=f'median: {median_loon:.2f}')

plt.xlabel("Loon")
plt.ylabel("Density")

plt.title("verdeling loon sector 1636")

plt.gca().spines['top'].set_visible(False)
plt.gca().spines['right'].set_visible(False)

plt.legend()


# %%

# Writing basic statistics and graphs to excel file

excel_path = 'RVU.xlsx'
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

df_spolis_22['LOON_BB'] = df_spolis_22['SBASISLOON'] + df_spolis_22['SBIJZONDEREBELONING']
df_spolis_22['LOON_OW'] = df_spolis_22['SBASISLOON'] + df_spolis_22['SLNOWRK']
df_spolis_22['LOON_ALL'] = df_spolis_22['SBASISLOON'] + df_spolis_22['SBIJZONDEREBELONING'] + df_spolis_22['SLNOWRK']



#load VEHTAB 2022
fpath_veh_22 = r"G:\InkomenBestedingen\VEHTAB\VEH2022TABV1.sav" # set file path VEHTAB 2022
metadata_veh_22 = prt.read_sav(fpath_veh_22, metadataonly=True) # load metadata VEHTAB 2022
columns_to_load_veh_22 = ['RINPERSOONHKW', 'VEHW1000VERH', 'VEHW1121WONH'] #select columns to load from VEHTAB 2022
df_veh_22 = pd.read_spss(fpath_veh_22, usecols=columns_to_load_veh_22) #load VEHTAB 2022 data 


#load koppelbestand huishoudens -> personen 2022
fpath_koppel_veh_2022 = r"G:\InkomenBestedingen\VEHTAB\KOPPELPERSOONHUISHOUDEN2022V1.sav"
metadata_koppel_veh_22 = prt.read_sav(fpath_koppel_veh_2022, metadataonly=True)
columns_to_load__koppel_veh_22 = ['RINPERSOON', 'RINPERSOONHKW']
df_veh_koppel_22 = pd.read_spss(fpath_koppel_veh_2022, usecols=columns_to_load__koppel_veh_22)




# %%



