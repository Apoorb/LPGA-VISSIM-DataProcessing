# -*- coding: utf-8 -*-
"""
Created on Wed Aug  7 15:25:37 2019

@author: abibeka
"""
#0.0 Housekeeping. Clear variable space
from IPython import get_ipython  #run magic commands
ipython = get_ipython()
ipython.magic("reset -f")
ipython = get_ipython()


from bs4 import BeautifulSoup
import pandas as pd
from os import path
import geopy.distance
import numpy as np 
import subprocess #Need for opening Excel file from Python



def getDist(row):
    RowLatLong = tuple(row[['Lat','Long']].values.ravel())
    Names = ['LPGA@Tomoka','LPGA@I95-SB-Ramp','LPGA@I95-NB-Ramp','LPGA@Technology','LPGA@Willamson','LPGA@Clyde-Morris']
    Dist =[]
    for N1 in Names:
        Indx = DatLaLong.index[DatLaLong.Name==N1]
        Point = tuple(DatLaLong.loc[Indx,['Lat','Long']].values.ravel())
        Dist.append(geopy.distance.distance(Point,RowLatLong).feet)
    return pd.Series(Dist,index=['D1','D2','D3','D4','D5','D6'])

def getDist_Fun2(row):
        RowLatLong = tuple(row[['Lat','Long']].values.ravel())
        NextLatLong = tuple(row[['LatShift','LongShift']].values.ravel())
        if(row["LatShift"]==row["LatShift"]):
            return geopy.distance.distance(NextLatLong,RowLatLong).feet
        else: return 0
        
def GetSummaryData(RawData, DatLatLong, Runs=[5,7,9,11,13],Dir='EB'):
    
    if(Dir == 'EB'):
        #EB Heading = 60 Deg  #Runs Based on ZHB File
        EB_Mask = (RawData.Run.isin(Runs)) & (RawData.Heading.between(30,100))
        DataSub = RawData[EB_Mask]
    else: 
        WB_Mask = (RawData.Run.isin(Runs)) & (~RawData.Heading.between(30,100))
        DataSub = RawData[WB_Mask]

        
    DataSub[['D1','D2','D3','D4','D5','D6']] = DataSub[['Lat','Long']].apply(getDist,axis=1)
    DataSub.reset_index(inplace=True)
    DataSub['ClosestInt'] = np.nan
    for i in ['D1','D2','D3','D4','D5','D6']:
        Indx = DataSub.groupby('Run')[i].idxmin() #Find the location of D1, D2......D6
        DataSub.loc[Indx,'ClosestInt'] = i 
        
    DataSub.sort_values(['Run','Date_Time'],inplace=True) # Get sorted to shift columns
    DataSub[['DateTShift','LatShift','LongShift']] = DataSub.groupby('Run')['Date_Time','Lat','Long'].shift(-1)
    
    # Get the distance, speed and time just based on end points 
    DataBoundary = DataSub[~DataSub.ClosestInt.isna()]
    DataBoundary = DataBoundary[['Run','ClosestInt','Date_Time','Lat','Long']]
    DataBoundary = DataBoundary.sort_values(['Run','Date_Time']).reset_index(drop=True)
    CorSumDataBoundary = DataBoundary[DataBoundary.ClosestInt.isin(['D1','D6'])].reset_index(drop=True)
    DataBoundary[['DateTimeNextInt','LatShift','LongShift']] = DataBoundary.groupby('Run')['Date_Time','Lat','Long'].shift(-1)
    DataBoundary = DataBoundary[~DataBoundary.LatShift.isna()]
    DataBoundary.loc[:,"EndPointTimeDiff"] = (DataBoundary.DateTimeNextInt-DataBoundary.Date_Time).dt.total_seconds()
    DataBoundary.loc[:,'DistBtwEndPnt'] = DataBoundary[['Lat','Long','LatShift','LongShift']].apply(getDist_Fun2,axis=1)
    DataBoundary = DataBoundary[DataBoundary.DistBtwEndPnt>=900] #We know distance between each intersection is at least 900 ft.
    # Get the distance, speed and time just based on end points     
    CorSumDataBoundary[['DateTimeNextInt','LatShift','LongShift']] = CorSumDataBoundary.groupby('Run')['Date_Time','Lat','Long'].shift(-1)
    CorSumDataBoundary = CorSumDataBoundary[~CorSumDataBoundary.LatShift.isna()]
    CorSumDataBoundary.loc[:,"EndPointTimeDiff"] = (CorSumDataBoundary.DateTimeNextInt-CorSumDataBoundary.Date_Time).dt.total_seconds()
    CorSumDataBoundary.loc[:,'DistBtwEndPnt'] = CorSumDataBoundary[['Lat','Long','LatShift','LongShift']].apply(getDist_Fun2,axis=1)
    CorSumDataBoundary = CorSumDataBoundary.drop(columns='ClosestInt')
    
    # Subset the data and then use forward fill to set correct labels
    ListDa =[]    
    for key,group_df in DataSub.groupby('Run'):
        RefDT1 = group_df.Date_Time[group_df.ClosestInt == 'D1'].values[0]
        RefDT2 = group_df.Date_Time[group_df.ClosestInt == 'D6'].values[0]
        if(Dir == 'EB'):
            mask  = (group_df.Date_Time>=RefDT1) & (group_df.Date_Time<=RefDT2)
        else:
            mask  = (group_df.Date_Time<=RefDT1) & (group_df.Date_Time>=RefDT2)
        tempDat = group_df[mask]
        tempDat['ClosestInt'] = tempDat.ClosestInt.fillna(method='ffill')
        ListDa.append(tempDat)
    CleanDat = pd.concat(ListDa)
    CleanDat.reset_index(drop=True,inplace=True)
    CleanDat.loc[:,"TimeDiff"] = (CleanDat.DateTShift-CleanDat.Date_Time).dt.total_seconds()

    CleanDat.loc[:,'Speed_fps'] = CleanDat.Speed *1.47
    CleanDat.loc[:,'DistFromSpd'] = CleanDat.Speed_fps * CleanDat.TimeDiff
    CleanDat.loc[:,'DistBtwPnt'] = CleanDat[['Lat','Long','LatShift','LongShift']].apply(getDist_Fun2,axis=1)
    SumIntDat= CleanDat.groupby(['Run','ClosestInt']).agg({'DistFromSpd':'sum','DistBtwPnt':'sum','TimeDiff':'sum'})
    SumIntDat = SumIntDat.reset_index(drop=False)
    LeftVars = ['Run','ClosestInt']
    SumIntDat = pd.merge(DataBoundary,SumIntDat,left_on = LeftVars, right_on = LeftVars, how='left')
    FinIntDat = SumIntDat.groupby(['ClosestInt']).agg({'DistFromSpd':'mean','DistBtwPnt':'mean','TimeDiff':'mean'})
    FinIntDat.loc[:,'SMS_mph'] = FinIntDat.DistBtwPnt/ FinIntDat.TimeDiff/ 1.47
    SumCorDat = CleanDat.groupby(['Run']).agg({'DistFromSpd':'sum','DistBtwPnt':'sum','TimeDiff':'sum'})
    SumCorDat.loc[:,'SMS'] = SumCorDat.DistBtwPnt/SumCorDat.TimeDiff
    SumCorDat = SumCorDat.reset_index(drop=False)
    SumCorDat = pd.merge(CorSumDataBoundary,SumCorDat,left_on = 'Run', right_on = 'Run', how='left')
    CorridorSMS_mph = SumCorDat.DistBtwPnt.sum()/ SumCorDat.TimeDiff.sum()/ 1.47
    FinCorDat = pd.DataFrame({'MeanDist':[SumCorDat.DistBtwPnt.mean()], 'MeanTime': [SumCorDat.TimeDiff.mean()], 'SMS_mph': [CorridorSMS_mph]})
    return(SumIntDat,SumCorDat,FinIntDat,FinCorDat,CleanDat)
    
    
kml_file = path.join(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\LPGA-Intersections.kml')

with open(kml_file) as fp:
    soup = BeautifulSoup(fp,features='lxml')
#print(soup.prettify())

DatLaLong = pd.DataFrame()
for item in soup.find_all('placemark'):
    Name =item.find('name').string
    #Long = item.lookat.longitude.string
    #Lat = item.lookat.latitude.string
    Long,Lat= item.coordinates.string.split(',')[:2]
    tempDa = pd.DataFrame({'Name': [Name],'Lat': [Lat],'Long': [Long]})
    DatLaLong= pd.concat([DatLaLong,tempDa])
    print(Name,Lat,Long)
# DatLatLong has ref lat longs
DatLaLong.reset_index(inplace=True)
TTRunFi = path.join(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\LPGA Travel Time Data-AXB.xlsx')
    #Get the excel file in memory. Helps with iterating over sheet names
x1 = pd.ExcelFile(TTRunFi)
x1.sheet_names
AMTTRawDa= x1.parse('AM Raw',parse_dates=[['Date', 'Time']])
AMTTRawDa = AMTTRawDa[['Run','Date_Time','Speed','Heading','Lat','Long']]

PMTTRawDa= x1.parse('PM Raw',parse_dates=[['Date', 'Time']])
PMTTRawDa = PMTTRawDa[['Run','Date_Time','Speed','Heading','Lat','Long']]


Names = ['LPGA@Tomoka','LPGA@I95-SB-Ramp','LPGA@I95-NB-Ramp','LPGA@Technology','LPGA@Willamson','LPGA@Clyde-Morris']



FinData = {}

Debug = False
if (Debug): 
    #AM_EB Zach Coordinates
    DatLaLong.loc[DatLaLong.Name=='LPGA@Tomoka',['Lat','Long']] = ['29.217025','-81.109812']
    DatLaLong.loc[DatLaLong.Name=='LPGA@Clyde-Morris',['Lat','Long']] = ['29.227215','-81.085725']
    Run1 = np.arange(0,100)
else:
    Run1 = [5,7,9,11,13]
FinData['AM_EB'] =GetSummaryData(RawData=AMTTRawDa, DatLatLong = DatLaLong, Runs=Run1,Dir='EB')
FinData['AM_EB'][0]
FinData['AM_EB'][1]
FinData['AM_EB'][1].DistBtwPnt.sum() / FinData['AM_EB'][1].TimeDiff.sum() /1.47


if (Debug): 
    #AM_WB Zach Coordinates
    DatLaLong.loc[DatLaLong.Name=='LPGA@Clyde-Morris',['Lat','Long']] = ['29.22972','-81.08033']
    DatLaLong.loc[DatLaLong.Name=='LPGA@Tomoka',['Lat','Long']] = ['29.21753','-81.108792']
    Run1 = np.arange(0,100)
else:
    Run1 = [6,8,10,12,14]
FinData['AM_WB'] =GetSummaryData(RawData=AMTTRawDa, DatLatLong = DatLaLong, Runs=Run1 ,Dir='WB')
FinData['AM_WB'][0]
FinData['AM_WB'][1]
FinData['AM_WB'][1].DistBtwPnt.sum() / FinData['AM_WB'][1].TimeDiff.sum() /1.47



if (Debug): 
    #PM_EB Zach Coordinates
    DatLaLong.loc[DatLaLong.Name=='LPGA@Tomoka',['Lat','Long']] = ['29.217112','-81.109623']
    DatLaLong.loc[DatLaLong.Name=='LPGA@Clyde-Morris',['Lat','Long']] = ['29.227387','-81.085297']
    Run1 = [1,3,5,7,9,11,13]
else:
    Run1 = [5,7,9,11]
FinData['PM_EB'] =GetSummaryData(RawData=PMTTRawDa, DatLatLong = DatLaLong, Runs=Run1,Dir='EB')
FinData['PM_EB'][0]
FinData['PM_EB'][1]
FinData['PM_EB'][1].DistBtwPnt.sum() / FinData['PM_EB'][1].TimeDiff.sum() /1.47

if (Debug): 
    #PM_WB Zach Coordinates
    DatLaLong.loc[DatLaLong.Name=='LPGA@Clyde-Morris',['Lat','Long']] = ['29.229435','-81.080913']
    DatLaLong.loc[DatLaLong.Name=='LPGA@Tomoka',['Lat','Long']] = ['29.217562','-81.108732']
    Run1 = [2,4,6,8,10,12,14,16,17]
else:
    Run1 = [4,6,8,10,12]
FinData['PM_WB'] = GetSummaryData(RawData=PMTTRawDa, DatLatLong = DatLaLong, Runs=Run1,Dir='WB')

FinData['PM_WB'][1].DistBtwPnt.sum() / FinData['PM_WB'][1].TimeDiff.sum() /1.47
FinData['PM_WB'][0]
FinData['PM_WB'][1]
FinData['PM_WB'][1].DistBtwPnt.sum() / FinData['PM_WB'][1].TimeDiff.sum() /1.47
#Debug function manually
#RawData=AMTTRawDa; DatLatLong = DatLaLong; Runs=[3,5,7,9,11,13]; Dir='EB'




# VISSIM Travel Time Segments 
TTSeg= {1: 'EB',
            2: 'WB',
            3: 'EB LPGA (Tomoka Rd to I-95 SB Ramp)',
            4: 'EB LPGA (I-95 SB Ramp to I-95 NB Ramp)',
            5: 'EB LPGA (I-95 NB Ramp to Technology Blvd)',
            6: 'EB LPGA (Technology Blvd to Willamson Blvd)',
            7: 'EB LPGA (Willamson Blvd to Clyde-Morris Blvd)',
            8: 'WB LPGA (Clyde-Morris Blvd to Willamson Blvd)',
            9: 'WB LPGA (Willamson Blvd to Technology Blvd)',
            10:'WB LPGA (Technology Blvd to I-95 NB Ramp)',
            11:'WB LPGA (I-95 NB Ramp to I-95 SB Ramp)',
            12:'WB LPGA (I-95 SB Ramp to Tomoka Rd)'}

EBKeys = {'D1':3,
          'D2':4,
          'D3':5,
          'D4':6,
          'D5':7}
WBKeys = {'D6':8,
          'D5':9,
          'D4':10,
          'D3':11,
          'D2':12}



TTDebug = path.join(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\ComparisionValidat_Zach_TT.xlsx')
TTOut = path.join(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\TTData-by-Intersection.xlsx')

if (Debug): 
    writer=pd.ExcelWriter(TTDebug)
    for key in FinData.keys():
        FinData[key][3].to_excel(writer,key)
        startrow1 = writer.sheets[key].max_row
        FinData[key][1].to_excel(writer,key, startrow = startrow1+2)
    writer.save()
else:
    writer=pd.ExcelWriter(TTOut)
    for key in FinData.keys():
        dat = FinData[key][2] 
        if ((key == 'AM_EB') |(key == 'PM_EB')):
            dat= dat.rename(index=EBKeys).reset_index()
        else:
            dat= dat.rename(index=WBKeys).reset_index()
        dat.loc[:,'SegName'] = dat.ClosestInt.apply(lambda x:TTSeg[x])
        
        dat.to_excel(writer,key)
        startrow1 = writer.sheets[key].max_row
        FinData[key][0].to_excel(writer,key, startrow = startrow1+2)
        startrow1 = writer.sheets[key].max_row
        FinData[key][3].to_excel(writer,key, startrow = startrow1+2)
        startrow1 = writer.sheets[key].max_row
        FinData[key][1].to_excel(writer,key, startrow = startrow1+2)
    writer.save()
 
    
# Open excel file from python 
subprocess.Popen([TTOut],shell=True)  

subprocess.Popen([TTDebug],shell=True)  

subprocess.Popen([TTRunFi],shell=True)  


