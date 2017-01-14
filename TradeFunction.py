# -*- coding: utf-8 -*-
"""
Created on Tue Dec 13 14:53:48 2016

@author: Tayyab Rashid
"""

import sqlite3
import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter
import datetime
conn = sqlite3.connect('Shares2.sqlite3')
TradePool = pd.DataFrame(columns=('Date','Symbol','Price', 'Amount', 'OrderType' , 'Stoploss', 'TakeProfit','UseSL','UseTP'))
HistoricalTradePool = pd.DataFrame(columns=('Date','Symbol','Price', 'Amount', 'OrderType' , 'Stoploss', 'TakeProfit', 'ClosingPrice', 'Profit'))
PendingBuffer = pd.DataFrame(columns=('Symbol','Price', 'Amount', 'OrderType' , 'Stoploss', 'TakeProfit', 'ClosingPrice', 'Profit'))
Price=0
Profit=0
CumulativeProfit=pd.DataFrame({'TradeNumber':[0],'Profit':[0],'CumulativeProfit':[0]})

def ImportFinancialData():
    global EPS_Date
    data = pd.read_excel('EPS.xlsx', 'Ark1',skiprows=1)
    EPS_Date=pd.DataFrame(columns=('Date','Data'))
    for i in sorted(range(0,len(data)),reverse=True):
        EPS_Date.loc[len(EPS_Date)+1]=((data.iloc[i,0]).date(),(data.iloc[i,1]))
        
      


def ImportData(Symbol):
    global AllData
    global Date
    global Open
    global High
    global Low
    global Close
    global Volum
    global PriceData
    
    sql="SELECT* from Symbol"
    sql1=sql.replace("Symbol",Symbol)
    AllData = pd.read_sql_query(sql1,con=conn)
    Date=AllData.iloc[0:len(AllData),0]
    Open=AllData.iloc[0:len(AllData),1]
    High=AllData.iloc[0:len(AllData),2]
    Low=AllData.iloc[0:len(AllData),3]
    Close=AllData.iloc[0:len(AllData),4]
    Volum=AllData.iloc[0:len(AllData),5]
    PriceData=AllData.iloc[0:len(AllData),6]
    for d in range(0,len(Date)):
        Date[d]=datetime.datetime.strptime(Date[d], '%Y-%m-%d %H:%M:%S').date()
    a=(print(min(Date),max(Date)))
    return a



    
def OnBar():
    global Profit
    global CumulativeProfit
    global Price
    global TradePool
    global count
    global EPS_Date
    global s
    count=0
    ImportFinancialData()
    
    for j in range(0,len(PriceData)-20):
        
        
        
        
        Price=PriceData.iloc[j]
        s=j
        
       
        
        for k in range(7,len(EPS_Date)):
            if Date[s]>=EPS_Date.iloc[k,0]:
                
                EPS_Date = EPS_Date.drop([k])
                EPS_Date = EPS_Date.reset_index(drop=True)
                print("Bastard")
                count=count+1
                OrderSend(Security="MSFT",OrderType="BUYLIMIT",OrderPrice=Price-Price*0.01,AmountSize=10000,SL=5 ,TP=5)
                break
                #if (EPS_Date.iloc[k-1,1]+EPS_Date.iloc[k-2,1]+EPS_Date.iloc[k-3,1])/3>EPS_Date.iloc[k,1]:
                    
      
        
        
        CheckOpenTrade()
        if s==len(PriceData)-1:
            for i in reversed(range(0,len(TradePool))):
                if TradePool.iloc[i,3]=="OP_BUY":
                    HistoricalTradePool.loc[len(HistoricalTradePool)+1]=[TradePool.iloc[i,0],TradePool.iloc[i,1],TradePool.iloc[i,2] , TradePool.iloc[i,3], TradePool.iloc[i,4], TradePool.iloc[i,5], TradePool.iloc[i,6], Price, round((Price-TradePool.iloc[i,2])*(TradePool.iloc[i,3]/TradePool.iloc[i,2]),0)]                    
                if TradePool.iloc[i,3]=="OP_SELL":
                    HistoricalTradePool.loc[len(HistoricalTradePool)+1]=[TradePool.iloc[i,0],TradePool.iloc[i,1],TradePool.iloc[i,2] , TradePool.iloc[i,3], TradePool.iloc[i,4], TradePool.iloc[i,5], TradePool.iloc[i,6], Price, round((TradePool.iloc[i,2]-Price)*(TradePool.iloc[i,3]/TradePool.iloc[i,2]),0)] 
                TradePool=TradePool.drop([i])
                TradePool=TradePool.reset_index(drop=True)
                
               
    Profit=HistoricalTradePool['Profit']
    for i in range(1,len(Profit)):
        CumulativeProfit.loc[len(CumulativeProfit)+1]=[round(CumulativeProfit.iloc[i-1,0]+Profit[i],0),round(Profit[i],0),i]
        
    a=(round(Profit.sum(),0),round(Profit.describe(),0),plt.plot(CumulativeProfit.iloc[:,0], linewidth=2),count)
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    EventStudy() 
    workbook = xlsxwriter.Workbook('TradeData.xlsx')
    workbook.add_worksheet('HistoricalTrades')
    workbook.add_worksheet('CumulativePrft')
    workbook.add_worksheet('Eventstudy')
    writer = pd.ExcelWriter('TradeData.xlsx', engine='xlsxwriter')
    HistoricalTradePool.to_excel(writer, sheet_name='HistoricalTrades')
    CumulativeProfit.to_excel(writer, sheet_name='CumulativePrft')
    df.to_excel(writer, sheet_name='Eventstudy')
    
    
    return a

def OrderSend(Security, OrderType, OrderPrice=0, AmountSize=0, SL=0, TP=0):
    global Price
    global TradePool
    Security1=Security;
    OrderType=OrderType;
    UseSL=False;
    UseTP=False;
    if(AmountSize>0):
        Amount=AmountSize
        if(OrderType=="OP_BUY" or OrderType=="OP_SELL"):
              OrderPrice=Price
              if(OrderType=="OP_BUY"):
                  SL2=Price-(Price*SL/100)
                  TP2=Price+(Price*TP/100)
                  if(SL!=0 and SL2 < OrderPrice):
                      UseSL=True
                  if(TP!=0 and TP2> OrderPrice):
                      UseTP=True
                  TradePool.loc[len(TradePool)+1]=[Date.loc[s],Security1, OrderPrice, Amount, OrderType, SL2, TP2, UseSL, UseTP];
              if(OrderType=="OP_SELL"):
                  SL2=Price+(Price*SL/100)
                  TP2=Price-(Price*TP/100)
                  if(SL!=0 and SL2>OrderPrice):
                      UseSL=True
                  if(TP!=0 and TP2< OrderPrice):
                      UseTP=True
                  TradePool.loc[len(TradePool)+1]=[Date.loc[s],Security1, OrderPrice, Amount, OrderType, SL2, TP2, UseSL, UseTP]
        if(OrderType=="BUYLIMIT"):
            if(OrderPrice<Price):
                SL2=OrderPrice-(OrderPrice*SL/100)
                TP2=OrderPrice+(OrderPrice*TP/100)
                if(SL!=0 and SL2 < OrderPrice):
                    UseSL=True
                if(TP!=0 and TP2> OrderPrice):
                    UseTP=True
                TradePool.loc[len(TradePool)+1]=[Date.loc[s],Security1, OrderPrice, Amount, OrderType, SL2, TP2, UseSL, UseTP];
            else:
                print("Wrong price")
        if(OrderType=="BUYSTOP"):
            if(OrderPrice>Price):
                SL2=OrderPrice-(OrderPrice*SL/100)
                TP2=OrderPrice+(OrderPrice*TP/100)
                if(SL!=0 and SL2 < OrderPrice):
                    UseSL=True
                if(TP!=0 and TP2> OrderPrice):
                    UseTP=True
                TradePool.loc[len(TradePool)+1]=[Date.loc[s],Security1, OrderPrice, Amount, OrderType, SL2, TP2, UseSL, UseTP];
            else:
                print("Wrong price")
            
        if(OrderType=="SELLLIMIT"):
            if(OrderPrice>Price):
                SL2=OrderPrice+(OrderPrice*SL/100)
                TP2=OrderPrice-(OrderPrice*TP/100)
                if(SL!=0 and SL2 < OrderPrice):
                    UseSL=True
                if(TP!=0 and TP2> OrderPrice):
                    UseTP=True
                TradePool.loc[len(TradePool)+1]=[Date.loc[s],Security1, OrderPrice, Amount, OrderType, SL2, TP2, UseSL, UseTP];
            else:
                print("Wrong price")
        if(OrderType=="SELLSTOP"):
            if(OrderPrice<Price):
                SL2=OrderPrice+(OrderPrice*SL/100)
                TP2=OrderPrice-(OrderPrice*TP/100)
                if(SL!=0 and SL2 < OrderPrice):
                    UseSL=True
                if(TP!=0 and TP2> OrderPrice):
                    UseTP=True
                TradePool.loc[len(TradePool)+1]=[Date.loc[s],Security1, OrderPrice, Amount, OrderType, SL2, TP2, UseSL, UseTP];
            else:
                print("Wrong price")
            
    else:
        print("Error amount")
    return
    
def CheckOpenTrade():
    global TradePool
    global HistoricalTradePool
    for i in reversed(range(0,len(TradePool))):
        TradePool=TradePool.reset_index(drop=True)  
        if TradePool.iloc[i,4]=="OP_BUY":
            if(TradePool.iloc[i,7]==True):
                if(Price < TradePool.iloc[i,5]):
                    HistoricalTradePool.loc[len(HistoricalTradePool)+1]=[TradePool.iloc[i,0],TradePool.iloc[i,1],TradePool.iloc[i,2] , TradePool.iloc[i,3], TradePool.iloc[i,4], TradePool.iloc[i,5], TradePool.iloc[i,6], Price, round((Price-TradePool.iloc[i,2])*(TradePool.iloc[i,3]/TradePool.iloc[i,2]),0)]
                    TradePool=TradePool.drop([i])
                    continue
            if (TradePool.iloc[i,8]==True):
                if Price > TradePool.iloc[i,5]:
                    HistoricalTradePool.loc[len(HistoricalTradePool)+1]=[TradePool.iloc[i,0],TradePool.iloc[i,1],TradePool.iloc[i,2] , TradePool.iloc[i,3], TradePool.iloc[i,4], TradePool.iloc[i,5], TradePool.iloc[i,6], Price, round((Price-TradePool.iloc[i,2])*(TradePool.iloc[i,3]/TradePool.iloc[i,2]),0)]
                    TradePool=TradePool.drop([i])  
                    continue
        if TradePool.iloc[i,4]=="OP_SELL":
            if(TradePool.iloc[i,7]==True):
                if(Price > TradePool.iloc[i,5]):
                    HistoricalTradePool.loc[len(HistoricalTradePool)+1]=[TradePool.iloc[i,0],TradePool.iloc[i,1],TradePool.iloc[i,2] , TradePool.iloc[i,3], TradePool.iloc[i,4], TradePool.iloc[i,5], TradePool.iloc[i,6], Price, round((TradePool.iloc[i,2]-Price)*(TradePool.iloc[i,3]/TradePool.iloc[i,2]),0)] 
                    TradePool=TradePool.drop([i]) 
                    continue
            if(TradePool.iloc[i,8]==True):
                if Price < TradePool.iloc[i,6]:
                    HistoricalTradePool.loc[len(HistoricalTradePool)+1]=[TradePool.iloc[i,0],TradePool.iloc[i,1],TradePool.iloc[i,2] , TradePool.iloc[i,3], TradePool.iloc[i,4], TradePool.iloc[i,5], TradePool.iloc[i,6], Price, round((TradePool.iloc[i,2]-Price)*(TradePool.iloc[i,3]/TradePool.iloc[i,2]),0)]
                    TradePool=TradePool.drop([i])
                    
        if(TradePool.iloc[i,4]=="BUYLIMIT"):
            if(TradePool.iloc[i,2]>Price):
                TradePool.iloc[i,4]="OP_BUY"
                TradePool.iloc[i,2]=Price
                continue
                
        if(TradePool.iloc[i,4]=="BUYSTOP"):
            if(TradePool.iloc[i,2]<Price):
                TradePool.iloc[i,4]="OP_BUY"
                TradePool.iloc[i,2]=Price
                continue
        if(TradePool.iloc[i,4]=="SELLSTOP"):
            if(TradePool.iloc[i,2]>Price):
                TradePool.iloc[i,4]="OP_SELL"
                TradePool.iloc[i,2]=Price
                continue
                
        if(TradePool.iloc[i,4]=="SELLLIMIT"):
            if(TradePool.iloc[i,2]<Price):
                TradePool.iloc[i,4]="OP_SELL"
                TradePool.iloc[i,2]=Price
                continue
    return
            
def TotalOrders(Time):
    Total=0
    if Time=="Current":
        Total=len(TradePool)
    if Time=="Historical":
        Total=len(HistoricalTradePool)
    return Total
        
    
def OrderSelect(index, Time):
    global OrderSymbol
    global OrderPrice
    global OrderAmount
    global OrderType
    global OrderStopLoss
    global OrderStopLoss
    global OrderClosingPrice
    global OrderProfit
    global OrderInd
    global OrderTradeTime
    
    if(Time=="Current"):
        if(len(TradePool)>=1):
            OrderTradeTime=TradePool.iloc[index,0]
            OrderSymbol=TradePool.iloc[index,1]
            OrderPrice= TradePool.iloc[index,2]
            OrderAmount= TradePool.iloc[index,3]
            OrderType= TradePool.iloc[index,4]
            OrderStopLoss=TradePool.iloc[index,5]
            OrderStopLoss=TradePool.iloc[index,6]
            OrderInd=index
            a=True
        else:
            a=False
       
    if(Time=="Historical"):
        if(len(HistoricalTradePool)>0):
            OrderTradeTime=HistoricalTradePool.iloc[index,0]
            OrderSymbol=HistoricalTradePool.iloc[index,1]
            OrderPrice= HistoricalTradePool.iloc[index,2]
            OrderAmount= HistoricalTradePool.iloc[index,3]
            OrderType= HistoricalTradePool.iloc[index,4]
            OrderStopLoss=HistoricalTradePool.iloc[index,5]
            OrderStopLoss=HistoricalTradePool.iloc[index,6]
            OrderClosingPrice=HistoricalTradePool.iloc[index,7]
            OrderProfit=HistoricalTradePool.iloc[index,8]
            OrderInd=index
            a=True
        else:
            a=False
    return a

def OrderClose(OrderIndex=0):
    global OrderInd
    global TradePool
    global  HistoricalTradePool
    OrderIndex=OrderInd
    if TradePool.iloc[OrderIndex,3]=="OP_BUY":
        HistoricalTradePool.loc[len(HistoricalTradePool)+1]=[TradePool.iloc[OrderIndex,0],TradePool.iloc[OrderIndex,1],TradePool.iloc[OrderIndex,2] , TradePool.iloc[OrderIndex,3], TradePool.iloc[OrderIndex,4], TradePool.iloc[OrderIndex,5], TradePool.iloc[OrderIndex,6], Price, round((Price-TradePool.iloc[OrderIndex,2])*(TradePool.iloc[OrderIndex,3]/TradePool.iloc[OrderIndex,2]),0)]                    
    if TradePool.iloc[OrderIndex,3]=="OP_SELL":
        HistoricalTradePool.loc[len(HistoricalTradePool)+1]=[TradePool.iloc[OrderIndex,0],TradePool.iloc[OrderIndex,1],TradePool.iloc[OrderIndex,2] , TradePool.iloc[OrderIndex,3], TradePool.iloc[OrderIndex,4], TradePool.iloc[OrderIndex,5], TradePool.iloc[OrderIndex,6], Price, round((TradePool.iloc[OrderIndex,2]-Price)*(TradePool.iloc[OrderIndex,3]/TradePool.iloc[OrderIndex,2]),0)] 
    TradePool=TradePool.drop([OrderIndex])
    TradePool=TradePool.reset_index(drop=True)
    
            
def EventStudy():
    global df
    df = pd.DataFrame(index=range(20))
    for r in range(0,len(HistoricalTradePool)):
        TradeDate=HistoricalTradePool.iloc[r,0]
        for t in range(0,len(Date)):
            if(Date[t]==TradeDate):
                df[r]=[(PriceData[t+1]-PriceData[t])/PriceData[t]*100,(PriceData[t+2]-PriceData[t])/PriceData[t]*100,(PriceData[t+3]-PriceData[t])/PriceData[t]*100,(PriceData[t+4]-PriceData[t])/PriceData[t]*100,
(PriceData[t+5]-PriceData[t])/PriceData[t]*100, (PriceData[t+6]-PriceData[t])/PriceData[t]*100, (PriceData[t+7]-PriceData[t])/PriceData[t]*100, (PriceData[t+8]-PriceData[t])/PriceData[t]*100, (PriceData[t+9]-PriceData[t])/PriceData[t]*100,
(PriceData[t+10]-PriceData[t])/PriceData[t]*100, (PriceData[t+11]-PriceData[t])/PriceData[t]*100, (PriceData[t+12]-PriceData[t])/PriceData[t]*100, (PriceData[t+13]-PriceData[t])/PriceData[t]*100, (PriceData[t+14]-PriceData[t])/PriceData[t]*100, (PriceData[t+15]-PriceData[t])/PriceData[t]*100,
(PriceData[t+16]-PriceData[t])/PriceData[t]*100, (PriceData[t+17]-PriceData[t])/PriceData[t]*100, (PriceData[t+18]-PriceData[t])/PriceData[t]*100, (PriceData[t+19]-PriceData[t])/PriceData[t]*100, (PriceData[t+20]-PriceData[t])/PriceData[t]*100]
    
    
    plt.plot(df)
    
    
    
    
    
            
            