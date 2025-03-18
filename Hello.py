
import yfinance as yf
import yahoo_fin.stock_info as si
#import yahoo_finance as yfin
#import googlefinance as gfin
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
    link ='https://www.nikkoam.com/files/etf/_shared/xls/portfolio/'+FundCode+'_'+YearMonth+'.xls'
    #link ='https://www.dropbox.com/scl/fi/l7nda0z6loqo0gqvk6ums/2239_202401.xls?rlkey=khoxushwv1scbbfb3pkz5qvk5&dl=1'
    fileName, headers = urllib.request.urlretrieve(link)
    wb = xlrd.open_workbook(fileName, logfile=open(os.devnull, 'w'))
    
    fxJPY = yf.Ticker('JPY=X').history(period='1d', interval='1m')['Close'].iloc[-1] 
    futuresdf = pd.DataFrame([yf.Ticker("ESM25.CME").info])
    futLastPrice = yf.Ticker("ESM25.CME").fast_info['last_price'] #(futuresdf['ask'].iloc[0]+futuresdf['bid'].iloc[0])/2
    
    FundPositions = pd.read_excel(wb,skiprows= range(1, 13),skipfooter=3,header=1,engine='xlrd')
    futPositions = FundPositions[FundPositions.Category=="Future"]
    curFutPosition = futPositions['Value(Local)'].sum() / 5 /futPositions['Price'].mean()
    
    fundData = pd.read_excel(wb,header=None,usecols="A,C",skiprows=1, nrows=10,engine='xlrd').T
    fundData.columns = fundData.iloc[0]
    fundData = fundData[1:]
    fundData = fundData.infer_objects()
    futPctChange = futLastPrice/futPositions['Price'].mean()-1
    date2239 = pd.read_excel(wb,header=None,usecols="A",skiprows=0, nrows=1,engine='xlrd')
    nav = fundData['AUM*1'].iloc[0]
    LevRatio = 2
    
    cfFactor=(cF2239+nav)/nav
    
    targetPosition = LevRatio*nav*(1+LevRatio*futPctChange)/(fxJPY*5*futLastPrice)*cfFactor


    prevInvRatio2239 = futPositions['Value(JPY)'].sum()/nav
    targetTrade = targetPosition - curFutPosition
    liveFundWeight = curFutPosition / targetPosition * LevRatio
    st.title('Live')
    st.write(FundCode +':     '+ '{:+.2%}'.format(liveFundWeight)+' &nbsp; &nbsp;' + '{:.1f}'.format(targetTrade) +' Micros &nbsp; ')
    temp2239Pos = curFutPosition
    
    FundCode = "2240"
    #cF2240 = 0

    YearMonth = date.today().strftime('%Y%m')
    link ='https://www.nikkoam.com/files/etf/_shared/xls/portfolio/'+FundCode+'_'+YearMonth+'.xls'
    #link = 'https://www.dropbox.com/scl/fi/c89h5c99239p0itd5zp79/2240_202401.xls?rlkey=4vc3vk9jgddtjv5dted4heuad&dl=1'
    fileName, headers = urllib.request.urlretrieve(link)
    wb = xlrd.open_workbook(fileName, logfile=open(os.devnull, 'w'))

    fundData = pd.read_excel(wb,header=None,usecols="A,C",skiprows=1, nrows=10,engine='xlrd').T
    fundData.columns = fundData.iloc[0]
    fundData = fundData[1:]
    fundData = fundData.infer_objects()

    FundPositions = pd.read_excel(wb,skiprows= range(1, 13),skipfooter=3,header=1,engine='xlrd')
    futPositions = FundPositions[FundPositions.Category=="Future"]
    curFutPosition = futPositions['Value(Local)'].sum() / 5 /futPositions['Price'].mean()

    nav = fundData['AUM*1'].iloc[0]
    LevRatio = -1

    futPctChange = futLastPrice/futPositions['Price'].mean()-1
    
    cfFactor=(cF2240+nav)/nav
    
    targetPosition = LevRatio*nav*(1+LevRatio*futPctChange)/(fxJPY*5*futLastPrice)*cfFactor


    prevInvRatio2240 = futPositions['Value(JPY)'].sum()/nav
    targetTrade = targetPosition - curFutPosition
    liveFundWeight = curFutPosition / targetPosition * LevRatio
 
    st.write(FundCode +':     '+ '{:+.2%}'.format(liveFundWeight)+' &nbsp; &nbsp;' + '{:.1f}'.format(targetTrade) +' Micros &nbsp; ')
    st.markdown('---')
    st.write('Futures: '+'{:.7}'.format(futLastPrice)  +' &nbsp; &nbsp;'+'{:+.2%}'.format(futPctChange))
    st.write('JPY: '+'{:.5}'.format(fxJPY) +' &nbsp; &nbsp;'+ '{:+.2%}'.format(futPositions['FX Rate'].mean()/fxJPY-1))
    #st.write('Current Position - &nbsp; 2239: '+ '{:.0f}'.format(temp2239Pos)+' &nbsp; &nbsp;' +' 2240: '+ '{:.0f}'.format(curFutPosition))

    st.markdown('---')
    st.write(date2239.at[0,0] + ' NAV Check')
    st.write('2239: '+'{:+.2%}'.format(prevInvRatio2239) +'&nbsp; | &nbsp;' + '{:.0f}'.format(temp2239Pos))
    st.write('2240: '+'{:+.2%}'.format(prevInvRatio2240)+'&nbsp; | &nbsp;' + '{:.0f}'.format(curFutPosition))

    
if __name__ == "__main__":

    cF2239 = st.number_input('2239 CF:',step=10000000.00, format="%f")
    cF2240 = st.number_input('2240 CF:',step=10000000.00,format="%f")
    calcOutputs(cF2239,cF2240)
