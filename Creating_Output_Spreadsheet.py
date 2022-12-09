'''
Creating the output spreadsheet
'''
import pandas as pd
import numpy as np
import openpyxl as xl
import datetime
from datetime import datetime as dt


GS_df = pd.read_excel('Data_Formatted/EPS_CHANGE_20221028_GS_FORMATTED.xlsx')
MS_df = pd.read_excel('Data_Formatted/EPS_CHANGES_20221028_MS_FORMATTED.xlsx')
JPMCAZ_df = pd.read_excel('Data_FORMATTED/EPS_CHANGE_20221028_JPMCAZ_FORMATTED.xlsx')
INTERMONTE_df = pd.read_excel('Data_Formatted/EPS_CHANGES_20221028_INTERMONTE_FORMATTED.xlsx')

#Combine input excel sheets into one dataframe
Combined_df = pd.concat([GS_df, MS_df, JPMCAZ_df, INTERMONTE_df])
#Drop first column
Combined_df.drop([Combined_df.columns[0]], axis=1, inplace=True)

#Replace missing valyes with 'N/A'
Combined_df = Combined_df.replace(np.nan, 'N/A')
#Sort rows alphabetically by Ticker
Combined_df = Combined_df.sort_values('Ticker')

Combined_df.reset_index(inplace=True, drop=True)
#Emptying duplicate ticker cells
Tickers_set = set()
for i in range(len(Combined_df)):
    if Combined_df.iloc[i]['Ticker'] not in Tickers_set:
        Tickers_set.add(Combined_df.iloc[i]['Ticker'])
    else:
        Combined_df.at[i,'Ticker'] = np.nan
        Combined_df.at[i,'Company Name'] = np.nan

# Formatting <><><><><><><><><><><><><><><><><><><><><><><><>

# Swapping columns to match example output
Combined_df = Combined_df[['Ticker', 'Company Name', 'Broker', 'currency', 'y1_EPS chg', 'y2_EPS chg', 'y3_EPS chg',
'y1 new EPS difference to consensus', 'y2 new EPS difference to consensus', 'y2 new EPS difference to consensus'
, 'y1_PE', 'y2_PE', 'y3_PE', 'comment', 'y1 new EPS', 'y2 new EPS', 'y3 new EPS',
'y1_sales chg', 'y2_sales chg', 'y3_sales chg', 'y1_EBIT chg', 'y2_EBIT chg', 'y3_EBIT chg']]



Combined_df.to_excel('Data_Formatted/OUTPUT.xlsx', sheet_name='Sheet1')


# print(Combined_df)
