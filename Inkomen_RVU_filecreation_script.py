# %%
# Import Packages
import pandas as pd
import statsmodels.api as sm
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import scipy as sp
import pyreadstat as prt

from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.writer.excel import ExcelWriter
from openpyxl import load_workbook
import io

# %%
#load SPOLISBUS 2020
fpath_spolis_20 = r"G:\Spolis\SPOLISBUS\2020\SPOLISBUS2020V5.sav" # set file path spolisbus 2020
metadata_spolis_20 = prt.read_sav(fpath_spolis_20, metadataonly=True) # load metadata spolisbus 2020
columns_to_load_spolisbus = ['RINPERSOON', 'SBASISLOON', 'SBASISUREN'] #select columns to load from spolisbus 2020
df_spolis_20 = pd.read_spss(fpath_spolis_20, usecols=columns_to_load_spolisbus) #load spolisbus data 2020

#load SPOLISBUS 2021
fpath_spolis_21 = r"G:\Spolis\SPOLISBUS\2021\SPOLISBUS2021V5.sav"# set file path spolisbus 2021
metadata_spolis_21 = prt.read_sav(fpath_spolis_21, metadataonly=True) # load metadata spolisbus 2021
columns_to_load_spolisbus = ['RINPERSOON', 'SBASISLOON', 'SBASISUREN'] #select columns to load from spolisbus 2021
df_spolis_21 = pd.read_spss(fpath_spolis_21, usecols=columns_to_load_spolisbus) #load spolisbus data 2021

#load SPOLISBUS 2022
fpath_spolis_22 = r"G:\Spolis\SPOLISBUS\2022\SPOLISBUS2022V5.sav" # set file path spolisbus 2022
metadata_spolis_22 = prt.read_sav(fpath_spolis_22, metadataonly=True) # load metadata spolisbus 2022
columns_to_load_spolisbus = ['RINPERSOON', 'SBASISLOON', 'SBASISUREN'] #select columns to load from spolisbus 2022
df_spolis_22= pd.read_spss(fpath_spolis_22, usecols=columns_to_load_spolisbus) #load spolisbus data 2022

#load SPOLISBUS 2023
fpath_spolis = r"G:\Spolis\SPOLISBUS\2023\SPOLISBUS2023V4.sav" # set file path spolisbus 2023
metadata_spolis_23 = prt.read_sav(fpath_spolis, metadataonly=True) # load metadata spolisbus 2023
columns_to_load_spolisbus = ['RINPERSOON', 'SBASISLOON', 'SBASISUREN'] #select columns to load from spolisbus 2023
df_spolis_23 = pd.read_spss(fpath_spolis, usecols=columns_to_load_spolisbus) #load spolisbus data 2023

# %%
#Load INPATAB
fpath_inpatab = "G:\InkomenBestedingen\INPATAB\INPA2022TABV2.sav" # set file path inpatab 2022
metadata_inpatab = prt.read_sav(fpath_inpatab, metadataonly=True) # load metadata inpatab 2022
columns_to_load_inpatab = ['RINPERSOON','INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB', 'INPPOSHHK'] #select columns to load from inpatab 2022
df_inpatab_22 = pd.read_spss (fpath_inpatab, usecols=columns_to_load_inpatab) #load inpatab 2022 data

# INPPERSBRUT is het bruto persoonlijk  inkomen (uit werk/eigen onderneming/ uitkeringe/toeslagen/etc.)
# INPPERSINK is het persoonlijk  inkomen (uit werk/eigen onderneming/ uitkeringe/toeslagen/etc.) met premies inkomensverzekering in mindering gebracht.
# INPPERSPRIM is het persoonlijk primair inkomen (inkomen uit werk en eigen onderneming)
# INPT1000WER is het loon van een werknemer/ inkomen uit arbeid
# INPT1020AMB is het loon van een ambtenaar/ inkomen uit arbeid ambtenaar
# INPPOSHHK is de positie in het huishouden (hoofdkostwinner etc.)

#Convert income columns to numeric (float)
df_inpatab_22[['INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB']] = df_inpatab_22[['INPPERSBRUT','INPPERSINK','INPPERSPRIM', 'INPT1000WER', 'INPT1020AMB']].apply(pd.to_numeric, errors='coerce')

# create UITWERK column
df_inpatab_22['UITWERK'] = df_inpatab_22['INPT1000WER'] + df_inpatab_22['INPT1020AMB']


# Create an income category variable

bins = [-float('inf'),0,5000,10000,15000,20000,25000,30000,35000,40000,45000,50000,55000,60000,65000,70000,75000,80000,85000,90000,95000,100000,105000,110000,115000,120000,125000,130000,135000,140000,145000,150000,155000,160000,165000,170000,175000,180000,185000,190000,195000,200000,205000,210000,215000,220000,225000,230000,235000,240000,245000,250000,255000,260000,265000,270000,275000,280000,285000,290000,295000,300000,float('inf')]
labels = ['negative', '0-4999','5000-9999','10000-14999','15000-19999','20000-24999','25000-29999','30000-34999','35000-39999','40000-44999','45000-49999','50000-54999','55000-59999','60000-64999','65000-69999','70000-74999','75000-79999','80000-84999','85000-89999','90000-94999','95000-99999',
          '100000-104999','105000-109999','110000-114999','115000-119999','120000-124999','125000-129999','130000-134999','135000-139999','140000-144999','145000-149999','150000-154999','155000-159999','160000-164999','165000-169999','170000-174999','175000-179999','180000-184999','185000-189999','190000-194999','195000-199999',
          '200000-204999','205000-209999','210000-214999','215000-219999','220000-224999','225000-229999','230000-234999','235000-239999','240000-244999','245000-249999','250000-254999','255000-259999','260000-264999','265000-269999','270000-274999','275000-279999','280000-284999','285000-289999','290000-294999','295000-299999','more than 300000']

df_inpatab_22['INCOME_CLASS'] = pd.cut(df_inpatab_22['UITWERK'], bins=bins, labels=labels, right=False)



# %%
# For each unique ID number all registered wages in SPOLISBUS are summed for the different years
df_spolis_20_sum = df_spolis_20.groupby('RINPERSOON')[['SBASISLOON','SBASISUREN']].sum().reset_index()
df_spolis_21_sum = df_spolis_21.groupby('RINPERSOON')[['SBASISLOON','SBASISUREN']].sum().reset_index()
df_spolis_22_sum = df_spolis_22.groupby('RINPERSOON')[['SBASISLOON','SBASISUREN']].sum().reset_index()
df_spolis_23_sum = df_spolis_23.groupby('RINPERSOON')[['SBASISLOON','SBASISUREN']].sum().reset_index()

# %%
# Add year column to each dataframe

df_spolis_20_sum['JAAR'] = 2020
df_spolis_21_sum['JAAR'] = 2021
df_spolis_22_sum['JAAR'] = 2022
df_spolis_23_sum['JAAR'] = 2023


# %%
#Concatenate to create SPOLIS 2020-2023
df_spolis_tot = pd.concat([df_spolis_20_sum, df_spolis_21_sum, df_spolis_22_sum, df_spolis_23_sum])

# Add a column with the hourly wage
df_spolis_tot['UURLOON'] = df_spolis_tot['SBASISLOON'] / df_spolis_tot['SBASISUREN']

# %%
# Sort by RINPERSOON and year.
df_spolis_tot_sort = df_spolis_tot.sort_values(by=['RINPERSOON','JAAR'])

# %%
#Filter for the second to last year that a person was in SPOLIS
df_spolis_tot_sort_2ndlast = df_spolis_tot_sort.groupby('RINPERSOON').nth(-2).reset_index()

# %%
#Keep the last year in SPOLIS
df_spolis_tot_sort_last = df_spolis_tot_sort.drop_duplicates(subset='RINPERSOON', keep='last')

# %%
df_spolis_tot_sort_last.head(30)

# %%
df_spolis_tot_sort_2ndlast.head(30)

# %%



