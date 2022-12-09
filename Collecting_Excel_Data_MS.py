'''
Collect Relevant Data from MS spreadsheet and store in df
'''

import pandas as pd
import numpy as np
import openpyxl as xl
import datetime
from datetime import datetime as dt
from Initialise import *

pd.set_option('display.max_rows', 3)
pd.set_option('display.max_columns', 10)


GS_dict = {'Company Name': 'Company Name', 'Ticker': 'BBG Ticker/ Currency', 'currency': 'BBG Ticker/ Currency',
'EPS_chg': 'EPS % change in JPM Estimate', 'new_EPS': 'EPS JPM (new)', 'new_EPS_difference_to_consensus': 'EPS % diff vs. consensus',
'PE': 'Cons PER (x)', 'EBIT_chg': 'JPMe EBIT', 'comment': 'Comments', 'FY': 'Forecast Year'}

MS_dict = {'Company Name': 'Company Name', 'Ticker': 'BBG Ticker/ Currency', 'currency': 'BBG Ticker/ Currency',
'EPS_chg': 'EPS % change in JPM Estimate', 'new_EPS': 'EPS JPM (new)', 'new_EPS_difference_to_consensus': 'EPS % diff vs. consensus',
'PE': 'Cons PER (x)', 'EBIT_chg': 'JPMe EBIT', 'comment': 'Comments', 'FY': 'Forecast Year'}

MS_df = pd.DataFrame(columns = Relevant_info_list)

MS_Data_df = pd.read_excel('Data/EPS_CHANGES_20221028_MS.xls', header = [0,1])


for j in range(0, 58, 2):
    Ticker_df = MS_Data_df.iloc[[j, j+1]]

    #Ensure dictionary is empty
    Ticker_dict.update((key, np.nan) for key in Ticker_dict)
    
    Ticker_dict['Broker'] = 'MS'
    Ticker_dict['Company Name'] = Ticker_df.iloc[0,0]
    Ticker_dict['Ticker'] = Ticker_df.iloc[0,1]
    Ticker_dict['currency'] = Ticker_df.iloc[1,1]
    Ticker_dict['comment'] = Ticker_df.iloc[0,8]

    for i in range(2):
        year = Ticker_df.iloc[i,2]
        year_number = year_dict.get(year)
        if year_number:
            Ticker_dict[f'{year_number}_PE'] = Ticker_df.iloc[i, 7]
            Ticker_dict[f'{year_number}_EPS chg'] = Ticker_df.iloc[i, 3]
            Ticker_dict[f'{year_number} new EPS difference to consensus'] = Ticker_df.iloc[i, 6]
            Ticker_dict[f'{year_number}_EBIT chg'] = Ticker_df.iloc[i, 10] #OP in MS sheet
            Ticker_dict[f'{year_number} new EPS'] = Ticker_df.iloc[i, 4] #Assuming MSe = Morgan Stanley estimate = new estimate
            Ticker_dict[f'{year_number}_sales chg'] = Ticker_df.iloc[i, 9]


    Ticker_df = pd.DataFrame([Ticker_dict])
    MS_df = pd.concat([MS_df, Ticker_df], ignore_index=True)

# print(MS _df)

MS_df.to_excel('Data_Formatted/EPS_CHANGES_20221028_MS_FORMATTED.xlsx', sheet_name='Sheet1')
