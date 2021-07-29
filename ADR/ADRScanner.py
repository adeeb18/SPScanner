# -*- coding: utf-8 -*-
"""
Created on Tue Jul 27 23:45:53 2021

@author: adeeb
"""
#Import modules
from pandas_datareader import data as pdr
from yahoo_fin import stock_info as si
from pandas import ExcelWriter
import yfinance as yf
import pandas as pd
import datetime
import time

#Declare Variables
tickers = si.tickers_sp500()
tickers = [item.replace(".", "-") for item in tickers] # Yahoo Finance uses dashes instead of dots
index_name = '^GSPC' # S&P 500
start_date = datetime.date.today() - datetime.timedelta(days=5)
end_date = datetime.date.today()
print(end_date)
print(start_date)
exportList = pd.DataFrame(columns=['Stock', "ADR"])

s = fr'C:\Users\adeeb\OneDrive\Documents\Adeeb\Stocks\HL_{end_date}.xlsx'
print(s)

HLDict = {}
ODict = {}

#Loop through all tickers in SP500
for ticker in tickers:
    #Create CSV for Stock
    df = pdr.get_data_yahoo(ticker, start_date, end_date)
    df.to_csv(fr'C:\Users\adeeb\OneDrive\Documents\Adeeb\Stocks\ADR\Stock Excel Dump\{ticker}.csv')
    
    #Calculate ADRs
    df['ADRHL'] = (df['High']-df['Low']) /df['Low']
    df['ADROH']=abs(df['High']-df['Open'])
    df['ADROL'] = abs(df['Low']-df['Open'])
    
    df['ADRO'] = (df[['ADROH','ADROL']].max(axis=1))/df['Open']
    
    stockADRHL = df['ADRHL'].mean()
    stockADRO = df['ADRO'].mean()
    
    #Assign to dictionaries
    HLDict[ticker] = stockADRHL
    ODict[ticker] = stockADRO
    
    print(ticker + " done");
    
#Sort HL
HLDictSorted = {}
sorted_keys = sorted(HLDict, key = HLDict.get)
for w in sorted_keys:
    HLDictSorted[w] = HLDict[w]

df = pd.DataFrame(HLDictSorted, index=[0])
df = (df.T)
df.to_excel(fr'C:\Users\adeeb\OneDrive\Documents\Adeeb\Stocks\ADR\HL\{end_date}.xlsx')

#Sort O
ODictSorted = {}
sorted_keys2 = sorted(ODict, key = ODict.get)
for w in sorted_keys:
    ODictSorted[w] = ODict[w]


df2 = pd.DataFrame(ODictSorted, index=[0])
df2 = (df2.T)
df2.to_excel(fr'C:\Users\adeeb\OneDrive\Documents\Adeeb\Stocks\ADR\O\{end_date}.xlsx')
