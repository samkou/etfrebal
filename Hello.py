
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


def calcOutputs(cF2239,cF2240):    

    FundCode = "2239"
    #cF2239 = 0
    
    YearMonth = date.today().strftime('%Y%m')
    #link ='https://www.nikkoam.com/files/etf/_shared/xls/portfolio/'+FundCode+'_'+YearMonth+'.xls'
    link ='https://www.dropbox.com/scl/fi/l7nda0z6loqo0gqvk6ums/2239_202401.xls?rlkey=khoxushwv1scbbfb3pkz5qvk5&dl=1'
    fileName, headers = urllib.request.urlretrieve(link)
    wb = xlrd.open_workbook(fileName, logfile=open(os.devnull, 'w'))

    fundData = pd.read_excel(wb,header=None,usecols="A,C",skiprows=1, nrows=10,engine='xlrd').T
    fundData.columns = fundData.iloc[0]
    fundData=fundData[1:]
    fundData = fundData.infer_objects()

    nav = fundData['AUM*1'].iloc[0]
    LevRatio = 2
    futuresdf = pd.DataFrame([yf.Ticker("ES=F").info])
    futLastPrice = yf.Ticker("ES=F").fast_info['last_price'] #(futuresdf['ask'].iloc[0]+futuresdf['bid'].iloc[0])/2
    futPctChange = futLastPrice/futuresdf['previousClose'].iloc[0]-1
    fxJPY = yf.Ticker("JPY=X").fast_info['last_price']
    cfFactor=(cF2239+nav)/nav

    targetPosition = LevRatio*nav*(1+LevRatio*futPctChange)/(fxJPY*5*futLastPrice)*cfFactor

    FundPositions = pd.read_excel(wb,skiprows= range(1, 13),skipfooter=3,header=1,engine='xlrd')
    futPositions = FundPositions[FundPositions.Category=="Future"]
    curFutPostion = futPositions['Value(Local)'].sum() / 5 /futPositions['Price'].mean()

    targetTrade = targetPosition - curFutPostion
    liveFundWeight = curFutPostion / targetPosition * LevRatio
    st.write(FundCode +':     '+ '{:+.2%}'.format(liveFundWeight)+' &nbsp; &nbsp;' + '{:.1f}'.format(targetTrade) +' Micros &nbsp; ')

    FundCode = "2240"
    #cF2240 = 0

    YearMonth = date.today().strftime('%Y%m')
    #link ='https://www.nikkoam.com/files/etf/_shared/xls/portfolio/'+FundCode+'_'+YearMonth+'.xls'
    link = 'https://www.dropbox.com/scl/fi/hsk6n7vc6clnimap09ay8/2240_202401.xls?rlkey=nrz4fy22kx1d6tb6pq07yw4wx&dl=1'
    fileName, headers = urllib.request.urlretrieve(link)
    wb = xlrd.open_workbook(fileName, logfile=open(os.devnull, 'w'))

    fundData = pd.read_excel(wb,header=None,usecols="A,C",skiprows=1, nrows=10,engine='xlrd').T
    fundData.columns = fundData.iloc[0]
    fundData=fundData[1:]
    fundData = fundData.infer_objects()

    nav = fundData['AUM*1'].iloc[0]
    LevRatio = -1
    futuresdf = pd.DataFrame([yf.Ticker("ES=F").info])
    futLastPrice = yf.Ticker("ES=F").fast_info['last_price'] #(futuresdf['ask'].iloc[0]+futuresdf['bid'].iloc[0])/2
    futPctChange = futLastPrice/futuresdf['previousClose'].iloc[0]-1
    fxJPY = yf.Ticker("JPY=X").fast_info['last_price']
    cfFactor=(cF2240+nav)/nav

    targetPosition = LevRatio*nav*(1+LevRatio*futPctChange)/(fxJPY*50*futLastPrice)*cfFactor

    FundPositions = pd.read_excel(wb,skiprows= range(1, 13),skipfooter=3,header=1,engine='xlrd')
    futPositions = FundPositions[FundPositions.Category=="Future"]
    curFutPostion = futPositions['Value(Local)'].sum() / 50 /futPositions['Price'].mean()

    targetTrade = targetPosition - curFutPostion
    liveFundWeight = curFutPostion / targetPosition * LevRatio
    st.write(FundCode +':     '+ '{:+.2%}'.format(liveFundWeight)+' &nbsp; &nbsp;' + '{:.1f}'.format(targetTrade) +' Minis &nbsp; ')
    st.write('Futures: '+'{:.6}'.format(futLastPrice)  +' &nbsp; &nbsp;'+'{:+.2%}'.format(futPctChange))
    st.write()
    st.write('JPY: '+'{:.5}'.format(fxJPY) +' &nbsp; &nbsp;'+ '{:+.2%}'.format(futPositions['FX Rate'].mean()/fxJPY-1))
    

if __name__ == "__main__":
    
    cF2239 = st.number_input('2239 CF:',step=100000000.00, format="%f")
    cF2240 = st.number_input('2240 CF:',step=100000000.00,format="%f")
    calcOutputs(cF2239,cF2240)
