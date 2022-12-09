'''
Collect Relevant Data from INTERMONTE spreadsheet and store in df

NB: All stocks in EUR and no comments
NB: Intermonte spreadsheet contained Factset XML Objects, hence the use of
openpyxl to open the sheet
'''

import pandas as pd
import numpy as np
import openpyxl as xl
import datetime
from datetime import datetime as dt
from Initialise import *

pd.set_option('display.max_rows', 5)
pd.set_option('display.max_columns', 30)

# INTERMONTE_dict = {'Company Name': 'Company ', 'Ticker': 'BBG Ticker', 'currency': 'BBG Ticker/ Currency',
# 'EPS_chg': 'EPS % change in JPM Estimate', 'new_EPS': 'EPS JPM (new)', 'new_EPS_difference_to_consensus': 'EPS % diff vs. consensus',
# 'PE': 'Cons PER (x)', 'EBIT_chg': 'JPMe EBIT', 'comment': 'Comments', 'FY': 'Forecast'}


INTERMONTE_df = pd.DataFrame(columns = Relevant_info_list)

wb = xl.load_workbook('Data/EPS_CHANGES_20221028_INTERMONTE.xlsx', data_only = True)
ws = wb.active
INTERMONTE_Data_df = pd.DataFrame(ws)


#Giving each cell its value, rather than an object
for i in range(len(INTERMONTE_Data_df)):
    for j in range(len(INTERMONTE_Data_df.columns)):
        INTERMONTE_Data_df.iloc[i,j] = INTERMONTE_Data_df.iloc[i,j].value

INTERMONTE_Data_df = INTERMONTE_Data_df.drop([0, 1])
INTERMONTE_Data_df = INTERMONTE_Data_df.reset_index(drop=True)


for j in range(2, 16, 2):

    Ticker_df = INTERMONTE_Data_df.iloc[[j, j+1]].reset_index(drop=True)

    #Ensure dictionary is empty
    Ticker_dict.update((key, np.nan) for key in Ticker_dict) 

    Ticker_dict['Broker'] = 'INTERMONTE'
    Ticker_dict['Company Name'] = Ticker_df.iloc[0,0]
    Ticker_dict['Ticker'] = Ticker_df.iloc[0,1]
    Ticker_dict['currency'] = Ticker_df.iloc[1,1]

    #Leave comment blank, will fill in blanks later with N/A

    for i in range(2):
        year = int(Ticker_df.iloc[i,5])
        year_number = year_dict.get(year)
        if year_number:
            Ticker_dict[f'{year_number}_PE'] = Ticker_df.iloc[i, 11]
            Ticker_dict[f'{year_number}_EPS chg'] = Ticker_df.iloc[i, 8]
            Ticker_dict[f'{year_number} new EPS difference to consensus'] = Ticker_df.iloc[i, 10]
            # Ticker_dict[f'{year_number}_EBIT chg'] = Ticker_df.iloc[i, 10] #NO EBIT DATA
            Ticker_dict[f'{year_number} new EPS'] = Ticker_df.iloc[i, 7]
            # Ticker_dict[f'{year_number}_sales chg'] = Ticker_df.iloc[i, 9] #NO SALES DATA


    Ticker_df = pd.DataFrame([Ticker_dict])
    INTERMONTE_df = pd.concat([INTERMONTE_df, Ticker_df], ignore_index=True)

INTERMONTE_df = INTERMONTE_df.replace('nm', None)


INTERMONTE_df.to_excel('Data_Formatted/EPS_CHANGES_20221028_INTERMONTE_FORMATTED.xlsx', sheet_name='Sheet1')
