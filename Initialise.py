'''
Initial Information Script
'''
import pandas as pd
import numpy as np
import openpyxl as xl
import datetime
from datetime import datetime as dt

Relevant_info_list = ['Company Name', 'Broker', 'Ticker', 'currency',
'y1_EPS chg', 'y1 new EPS', 'y1 new EPS difference to consensus',
'y2_EPS chg', 'y2 new EPS', 'y2 new EPS difference to consensus',
'y3_EPS chg', 'y3 new EPS', 'y3 new EPS difference to consensus',
'y1_PE','y2_PE','y3_PE',
'y1_sales chg',  'y1_EBIT chg',
'y2_sales chg',  'y2_EBIT chg',
'y3_sales chg',  'y3_EBIT chg',
'comment']

#Dictionary unique to each Ticker AND Broker containing the relevant information
Ticker_dict = dict.fromkeys(Relevant_info_list)
year_dict = {2023: 'y1', 2024: 'y2', 2025: 'y3'}
