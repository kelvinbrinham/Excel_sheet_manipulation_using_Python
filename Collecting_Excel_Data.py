import pandas as pd
import numpy as np
import openpyxl as xl
import datetime
from datetime import datetime as dt

pd.set_option('display.max_rows', 3)
pd.set_option('display.max_columns', 3)

#Make a common format data frame to store the relevant data for each broker
Relevant_info_list = ['Company Name', 'Broker', 'Ticker', 'currency',
'y1_EPS chg', 'y1 new EPS', 'y1 new EPS difference to consensus',
'y2_EPS chg', 'y2 new EPS', 'y2 new EPS difference to consensus',
'y3_EPS chg', 'y3 new EPS', 'y3 new EPS difference to consensus',
'y1_PE','y2_PE','y3_PE',
'y1_sales chg',  'y1_EBIT chg',
'y2_sales chg',  'y2_EBIT chg',
'y3_sales chg',  'y3_EBIT chg',
'comment']

JPMCAZ_dict = {'Company Name': 'Company Name', 'Ticker': 'BBG Ticker/ Currency', 'currency': 'BBG Ticker/ Currency',
'EPS_chg': 'EPS % change in JPM Estimate', 'new_EPS': 'EPS JPM (new)', 'new_EPS_difference_to_consensus': 'EPS % diff vs. consensus',
'PE': 'Cons PER (x)', 'EBIT_chg': 'JPMe EBIT', 'comment': '     Comments       ', 'FY': 'Forecast Year'}


JPMCAZ_df = pd.DataFrame(columns = Relevant_info_list)


JPMCAZ_Data_df = pd.read_excel('Data/EPS_CHANGE_20221028_JPMCAZ.xlsx', header = [3,4])
Tickers_list = [] #list of ticker dictionaries

Ticker_dict = dict.fromkeys(Relevant_info_list)
# print(Ticker_dict)

# print(JPMCAZ_Data_df)
#Try for one ticker to start with: Acerinox
Ticker_df = JPMCAZ_Data_df.iloc[[0, 1]]
# print(Ticker_df)
# Ticker_df = JPMCAZ_Data_df.iloc[[0+2, 1+2]].reset_index(drop=True)
#Company Name
# Ticker_dict['Company Name'] = Ticker_df['Company Name']['Unnamed: 0_level_1'][0]
Ticker_dict['Broker'] = 'JPMCAZ'
Ticker_dict['Company Name'] = Ticker_df['Company Name'].iloc[0][0]
Ticker_dict['Ticker'] = Ticker_df[JPMCAZ_dict.get('Ticker')].iloc[0][0]
Ticker_dict['currency'] = Ticker_df[JPMCAZ_dict.get('currency')].iloc[1][0]
Ticker_dict['comment'] = Ticker_df[JPMCAZ_dict.get('comment')].iloc[0][0]
year_dict = {2023: 'y1', 2024: 'y2', 2025: 'y3'}

for i in range(2):
    year = Ticker_df.iloc[i, 2]
    year_number = year_dict.get(year)
    if year_number:
        Ticker_dict[f'{year_number}_PE'] = Ticker_df.iloc[i, 8]
        Ticker_dict[f'{year_number}_EPS chg'] = Ticker_df.iloc[i, 3]
        Ticker_dict[f'{year_number} new EPS difference to consensus'] = Ticker_df.iloc[i, 6]
        Ticker_dict[f'{year_number}_EBIT chg'] = Ticker_df.iloc[i, 9]
        Ticker_dict[f'{year_number} new EPS'] = Ticker_df.iloc[i, 4]



print(Ticker_dict)
###
Ticker_df = pd.DataFrame([Ticker_dict])
JPMCAZ_df = pd.concat([JPMCAZ_df, Ticker_df], ignore_index=True)

# JPMCAZ_df.append(Ticker_dict, ignore_index=True)

# print(JPMCAZ_df)





# print(JPMCAZ_df)
JPMCAZ_df.to_excel('Data/OUTPUT.xlsx', sheet_name='Sheet1')

# Ticker_df.to_excel('Data/OUTPUT.xlsx', sheet_name='Sheet1')
