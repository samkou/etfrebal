# Copyright (c) Streamlit Inc. (2018-2022) Snowflake Inc. (2022)
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
import yfinance as yf
import yahoo_fin.stock_info as si
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
from datetime import date
import xlrd
import os
import requests
import xlrd
import urllib
import streamlit as st
from streamlit.logger import get_logger

LOGGER = get_logger(__name__)


def run():
    st.set_page_config(
        page_title="Hello",
        page_icon="👋",
    )

    FundCode = "2239"
    cF = 0

    YearMonth = date.today().strftime('%Y%m')
    link ='https://www.nikkoam.com/files/etf/_shared/xls/portfolio/'+FundCode+'_'+YearMonth+'.xls'
    fileName, headers = urllib.request.urlretrieve(link)
    wb = xlrd.open_workbook(fileName, logfile=open(os.devnull, 'w'))

    fundData = pd.read_excel(wb,header=None,usecols="A,C",skiprows=1, nrows=10,engine='xlrd').T
    fundData.columns = fundData.iloc[0]
    fundData=fundData[1:]
    fundData = fundData.infer_objects()

    NAV = fundData['AUM*1'].iloc[0]
    LevRatio = 2
    futuresdf = pd.DataFrame([yf.Ticker("ES=F").info])
    futLastPrice = (futuresdf['ask'].iloc[0]+futuresdf['bid'].iloc[0])/2
    futPctChange = futLastPrice/futuresdf['previousClose'].iloc[0]-1
    fxJPY = yf.Ticker("JPY=X").fast_info['last_price']
    cfFactor=(cF+NAV)/NAV

    TargetPosition = LevRatio*NAV*(1+LevRatio*futPctChange)/(fxJPY*5*futLastPrice)*cfFactor

    FundPositions = pd.read_excel(wb,skiprows= range(1, 13),skipfooter=3,header=1,engine='xlrd')
    futPositions = FundPositions[FundPositions.Category=="Future"]
    curFutPostion = futPositions['Value(Local)'].sum() / 5 /futPositions['Price'].mean()

    TargetTrade = TargetPosition - curFutPostion

    st.write(FundCode +':     '+str(TargetTrade.round(1))+' Micros')

    FundCode = "2240"
    cF = 0

    YearMonth = date.today().strftime('%Y%m')
    link ='https://www.nikkoam.com/files/etf/_shared/xls/portfolio/'+FundCode+'_'+YearMonth+'.xls'
    fileName, headers = urllib.request.urlretrieve(link)
    wb = xlrd.open_workbook(fileName, logfile=open(os.devnull, 'w'))

    fundData = pd.read_excel(wb,header=None,usecols="A,C",skiprows=1, nrows=10,engine='xlrd').T
    fundData.columns = fundData.iloc[0]
    fundData=fundData[1:]
    fundData = fundData.infer_objects()

    NAV = fundData['AUM*1'].iloc[0]
    LevRatio = -1
    futuresdf = pd.DataFrame([yf.Ticker("ES=F").info])
    futLastPrice = (futuresdf['ask'].iloc[0]+futuresdf['bid'].iloc[0])/2
    futPctChange = futLastPrice/futuresdf['previousClose'].iloc[0]-1
    fxJPY = yf.Ticker("JPY=X").fast_info['last_price']
    cfFactor=(cF+NAV)/NAV

    TargetPosition = LevRatio*NAV*(1+LevRatio*futPctChange)/(fxJPY*50*futLastPrice)*cfFactor

    FundPositions = pd.read_excel(wb,skiprows= range(1, 13),skipfooter=3,header=1,engine='xlrd')
    futPositions = FundPositions[FundPositions.Category=="Future"]
    curFutPostion = futPositions['Value(Local)'].sum() / 50 /futPositions['Price'].mean()

    TargetTrade = TargetPosition - curFutPostion

    st.write(FundCode +':     '+str(TargetTrade.round(1))+' Minis')



if __name__ == "__main__":
    run()
