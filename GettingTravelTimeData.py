# -*- coding: utf-8 -*-
"""
Created on Wed Aug  7 15:25:37 2019

@author: abibeka
"""
from bs4 import BeautifulSoup
import pandas as pd
from os import path
import geopy.distance
import numpy as np 

kml_file = path.join(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\LPGA-Intersections.kml')

with open(kml_file) as fp:
    soup = BeautifulSoup(fp)
#print(soup.prettify())

soup.placemark.find('name').string
soup.placemark.lookat.longitude.string
soup.placemark.lookat.latitude.string
soup.placemark.coordinates.string.split(',')[:2]
DatLaLong = pd.DataFrame()
for item in soup.find_all('placemark'):
    Name =item.find('name').string
    #Long = item.lookat.longitude.string
    #Lat = item.lookat.latitude.string
    Long,Lat= item.coordinates.string.split(',')[:2]
    tempDa = pd.DataFrame({'Name': [Name],'Lat': [Lat],'Long': [Long]})
    DatLaLong= pd.concat([DatLaLong,tempDa])
    print(Name,Lat,Long)


TTRunFi = path.join(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\LPGA Travel Time Data-AXB.xlsx')
    #Get the excel file in memory. Helps with iterating over sheet names
x1 = pd.ExcelFile(TTRunFi)
x1.sheet_names
AMTTRawDa= x1.parse('AM Raw',parse_dates=[['Date', 'Time']])
AMTTRawDa = AMTTRawDa[['Run','Date_Time','Speed','Heading','Lat','Long']]

#EB Heading = 60 Deg
EB_Mask = (AMTTRawDa.Run.isin([5,7,9,11,13])) & (AMTTRawDa.Heading.between(30,100))
EBAMTTRawDa = AMTTRawDa[EB_Mask]

# DatLatLong has ref lat longs
DatLaLong.reset_index(inplace=True)

def getDist(row):
    RowLatLong = tuple(row[['Lat','Long']].values.ravel())
    Indx1 = DatLaLong.index[DatLaLong.Name=='LPGA@Tomoka']
    Indx2 = DatLaLong.index[DatLaLong.Name=='LPGA@I95-SB-Ramp']
    Indx3 = DatLaLong.index[DatLaLong.Name=='LPGA@I95-NB-Ramp']
    Indx4 = DatLaLong.index[DatLaLong.Name=='LPGA@Technology']
    Indx5 = DatLaLong.index[DatLaLong.Name=='LPGA@Willamson']
    Indx6 = DatLaLong.index[DatLaLong.Name=='LPGA@Clyde-Morris']
    Point1 = tuple(DatLaLong.loc[Indx1,['Lat','Long']].values.ravel())
    Point2 = tuple(DatLaLong.loc[Indx2,['Lat','Long']].values.ravel())
    Point3 = tuple(DatLaLong.loc[Indx3,['Lat','Long']].values.ravel())
    Point4 = tuple(DatLaLong.loc[Indx4,['Lat','Long']].values.ravel())
    Point5 = tuple(DatLaLong.loc[Indx5,['Lat','Long']].values.ravel())
    Point6 = tuple(DatLaLong.loc[Indx6,['Lat','Long']].values.ravel())
    D1 = geopy.distance.distance(Point1,RowLatLong).feet
    D2 = geopy.distance.distance(Point2,RowLatLong).feet
    D3 = geopy.distance.distance(Point3,RowLatLong).feet
    D4 = geopy.distance.distance(Point4,RowLatLong).feet
    D5 = geopy.distance.distance(Point5,RowLatLong).feet
    D6 = geopy.distance.distance(Point6,RowLatLong).feet
    return pd.Series([D1,D2,D3,D4,D5,D6],index=['D1','D2','D3','D4','D5','D6'])
    
EBAMTTRawDa[['D1','D2','D3','D4','D5','D6']] = EBAMTTRawDa[['Lat','Long']].apply(getDist,axis=1)

EBAMTTRawDa.reset_index(inplace=True)
EBAMTTRawDa['ClosestInt'] = np.nan
for i in ['D1','D2','D3','D4','D5','D6']:
    Indx = EBAMTTRawDa.groupby('Run')[i].idxmin()
    EBAMTTRawDa.loc[Indx,'ClosestInt'] = i 
   

EBAMTTRawDa.sort_values(['Run','Date_Time'],inplace=True)


EBAMTTRawDa[['DateTShift','LatShift','LongShift']] = EBAMTTRawDa.groupby('Run')['Date_Time','Lat','Long'].shift(-1)


ListDa =[]
for key,group_df in EBAMTTRawDa.groupby('Run'):
    RefDT1 = group_df.Date_Time[group_df.ClosestInt == 'D1'].values[0]
    RefDT2 = group_df.Date_Time[group_df.ClosestInt == 'D6'].values[0]
    mask  = (group_df.Date_Time>=RefDT1) & (group_df.Date_Time<=RefDT2)
    tempDat = group_df[mask]
    tempDat['ClosestInt'] = tempDat.ClosestInt.fillna(method='ffill')
    ListDa.append(tempDat)

EBDat = pd.concat(ListDa)

EBDat.loc[:,"TimeDiff"] = (EBDat.DateTShift-EBDat.Date_Time).dt.total_seconds()


def getDist_Fun2(row):
    RowLatLong = tuple(row[['Lat','Long']].values.ravel())
    NextLatLong = tuple(row[['LatShift','LongShift']].values.ravel())
    if(row["LatShift"]==row["LatShift"]):
        return geopy.distance.distance(NextLatLong,RowLatLong).feet
    else: return 0

EBDat.loc[:,'Speed_fps'] = EBDat.Speed *1.47
EBDat.loc[:,'DistFromSpd'] = EBDat.Speed_fps * EBDat.TimeDiff
EBDat.loc[:,'DistBtwPnt'] = EBDat[['Lat','Long','LatShift','LongShift']].apply(getDist_Fun2,axis=1)



EBDat.groupby(['Run','ClosestInt']).agg({'DistFromSpd':'sum','DistBtwPnt':'sum','TimeDiff':'sum'})

EBDat.groupby(['Run']).agg({'DistFromSpd':'sum','DistBtwPnt':'sum','TimeDiff':'sum'})
