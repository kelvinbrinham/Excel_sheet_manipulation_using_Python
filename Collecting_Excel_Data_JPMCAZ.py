'''
Collect Relevant Data from JPMCAZ spreadsheet and store in df
'''

import pandas as pd
import numpy as np
import openpyxl as xl
import datetime
from datetime import datetime as dt
from Initialise import *

pd.set_option('display.max_rows', 3)
pd.set_option('display.max_columns', 3)



JPMCAZ_dict = {'Company Name': 'Company Name', 'Ticker': 'BBG Ticker/ Currency', 'currency': 'BBG Ticker/ Currency',
'EPS_chg': 'EPS % change in JPM Estimate', 'new_EPS': 'EPS JPM (new)', 'new_EPS_difference_to_consensus': 'EPS % diff vs. consensus',
'PE': 'Cons PER (x)', 'EBIT_chg': 'JPMe EBIT', 'comment': '     Comments       ', 'FY': 'Forecast Year'}


JPMCAZ_df = pd.DataFrame(columns = Relevant_info_list)


JPMCAZ_Data_df = pd.read_excel('Data/EPS_CHANGE_20221028_JPMCAZ.xlsx', header = [3,4])


for j in range(0, 60, 2):
    Ticker_df = JPMCAZ_Data_df.iloc[[j, j+1]]

    Ticker_dict['Broker'] = 'JPMCAZ'
    Ticker_dict['Company Name'] = Ticker_df['Company Name'].iloc[0][0]
    Ticker_dict['Ticker'] = Ticker_df[JPMCAZ_dict.get('Ticker')].iloc[0][0]
    Ticker_dict['currency'] = Ticker_df[JPMCAZ_dict.get('currency')].iloc[1][0]
    Ticker_dict['comment'] = Ticker_df[JPMCAZ_dict.get('comment')].iloc[0][0]


    for i in range(2):
        year = Ticker_df.iloc[i, 2]
        year_number = year_dict.get(year)
        if year_number:
            Ticker_dict[f'{year_number}_PE'] = Ticker_df.iloc[i, 8]
            Ticker_dict[f'{year_number}_EPS chg'] = Ticker_df.iloc[i, 3]
            Ticker_dict[f'{year_number} new EPS difference to consensus'] = Ticker_df.iloc[i, 6]
            Ticker_dict[f'{year_number}_EBIT chg'] = Ticker_df.iloc[i, 9]
            Ticker_dict[f'{year_number} new EPS'] = Ticker_df.iloc[i, 4]


    Ticker_df = pd.DataFrame([Ticker_dict])
    JPMCAZ_df = pd.concat([JPMCAZ_df, Ticker_df], ignore_index=True)



JPMCAZ_df.to_excel('Data_FORMATTED/EPS_CHANGE_20221028_JPMCAZ_FORMATTED.xlsx', sheet_name='Sheet1')
