'''
Collect Relevant Data from GS spreadsheet and store in df
'''

import pandas as pd
import numpy as np
import openpyxl as xl
import datetime
from datetime import datetime as dt
from Initialise import *

pd.set_option('display.max_rows', 3)
pd.set_option('display.max_columns', 3)


GS_dict = {'Company Name': 'Company Name', 'Ticker': 'BBG Ticker/ Currency', 'currency': 'BBG Ticker/ Currency',
'EPS_chg': 'EPS % change in JPM Estimate', 'new_EPS': 'EPS JPM (new)', 'new_EPS_difference_to_consensus': 'EPS % diff vs. consensus',
'PE': 'Cons PER (x)', 'EBIT_chg': 'JPMe EBIT', 'comment': 'Comments', 'FY': 'Forecast Year'}

GS_df = pd.DataFrame(columns = Relevant_info_list)

GS_Data_df = pd.read_excel('Data/EPS_CHANGE_20221028_GS.xlsx', header = [2,3])
GS_Data_df = GS_Data_df.drop(GS_Data_df.columns[[0]], axis=1)


for j in range(0, 70, 4):
    Ticker_df = GS_Data_df.iloc[[j + x for x in range(4)]]

    Ticker_dict['Broker'] = 'GS'
    Ticker_dict['Company Name'] = Ticker_df['Company Name'].iloc[0][0]
    Ticker_dict['Ticker'] = Ticker_df[GS_dict.get('Ticker')].iloc[0][0]
    Ticker_dict['currency'] = Ticker_df[GS_dict.get('currency')].iloc[1][0]
    Ticker_dict['comment'] = Ticker_df[GS_dict.get('comment')].iloc[0][0]

    for i in range(2):
        year = int(Ticker_df.iloc[i, 2][:4])
        year_number = year_dict.get(year)
        if year_number:
            Ticker_dict[f'{year_number}_PE'] = Ticker_df.iloc[i, 6]
            Ticker_dict[f'{year_number}_EPS chg'] = Ticker_df.iloc[i, 3]
            Ticker_dict[f'{year_number} new EPS difference to consensus'] = Ticker_df.iloc[i, 7]
            Ticker_dict[f'{year_number}_EBIT chg'] = Ticker_df.iloc[i, 10] #OP in GS sheet
            Ticker_dict[f'{year_number} new EPS'] = Ticker_df.iloc[i, 4]
            Ticker_dict[f'{year_number}_sales chg'] = Ticker_df.iloc[i, 9]


    Ticker_df = pd.DataFrame([Ticker_dict])
    GS_df = pd.concat([GS_df, Ticker_df], ignore_index=True)


GS_df.to_excel('Data_Formatted/EPS_CHANGE_20221028_GS_FORMATTED.xlsx', sheet_name='Sheet1')
