# -*- coding: utf-8 -*-
"""
Created on Mon Aug 19 13:47:01 2019

@author: abibeka
Read data from travel time files
"""

#0.0 Housekeeping. Clear variable space
from IPython import get_ipython  #run magic commands
ipython = get_ipython()
ipython.magic("reset -f")
ipython = get_ipython()



import os
import pandas as pd
import numpy as np
import glob
import datetime

def TTSegName(x):
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
    Nm = TTSeg[x]
    return Nm

# VISSIM File
#*********************************************************************************
PathToExist = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Existing'
ExistingPMfi = '20834_Existing_PM--C1C2C3C4C5C6C7_Vehicle Travel Time Results.att'
ExistingPMfi = os.path.join(PathToExist,ExistingPMfi)
ExistingAMfi ='20834_Existing_AM--C1C2aC3C4C5C6C7C8_Vehicle Travel Time Results.att'
ExistingAMfi = os.path.join(PathToExist,ExistingAMfi)

def PreProcessVissimTT(file = ExistingAMfi):
    PathToExist = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Existing'
    RefFile = '20834_Existing_PM--C1C2C3C4C5C6C7_Vehicle Travel Time Results.att'
    RefFile = os.path.join(PathToExist,RefFile)
    ExistingAMDat=pd.read_csv(file,sep =';',skiprows=17)
    ExistingAMDat.columns
    ExistingAMDat.rename(columns={'TRAVTM(ALL)':'VissimTT','VEHS(ALL)':'Veh'},inplace=True)
    mask=ExistingAMDat["$VEHICLETRAVELTIMEMEASUREMENTEVALUATION:SIMRUN"]=="AVG"
    ExistingAMDat = ExistingAMDat[mask]
    ExistingAMDat["TTSegNm"]=ExistingAMDat['VEHICLETRAVELTIMEMEASUREMENT'].apply(TTSegName)
    WB_TTSegs = ['WB LPGA (Clyde-Morris Blvd to Willamson Blvd)',
            'WB LPGA (Willamson Blvd to Technology Blvd)',
           'WB LPGA (Technology Blvd to I-95 NB Ramp)',
           'WB LPGA (I-95 NB Ramp to I-95 SB Ramp)',
            'WB LPGA (I-95 SB Ramp to Tomoka Rd)']
    if (file ==ExistingPMfi):
        mask1 = (ExistingAMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100'])) & (~ExistingAMDat.TTSegNm.isin(WB_TTSegs))
        #Can include '8100-9000'
        mask2 = (ExistingAMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100'])) & (ExistingAMDat.TTSegNm.isin(WB_TTSegs))
        mask = mask1 | mask2
    else:
        mask = ExistingAMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100']) 
    ExistingAMDat = ExistingAMDat[mask]
    ExistingAMDat.loc[:,'VissimTotalTT'] = ExistingAMDat.VissimTT * ExistingAMDat.Veh
    ExistingAMDat = ExistingAMDat.groupby(['VEHICLETRAVELTIMEMEASUREMENT','TTSegNm'])['VissimTotalTT','Veh'].sum().reset_index()
    ExistingAMDat.loc[:,'WeightedVissimTT'] = ExistingAMDat.VissimTotalTT/ExistingAMDat.Veh
    return(ExistingAMDat)

ExistingAMDat = PreProcessVissimTT(ExistingAMfi)
ExistingPMDat = PreProcessVissimTT(ExistingPMfi)

#*********************************************************************************
#TTRun Files
#*********************************************************************************
def ProcessObsData(TimePeriod,VissimDat):    
    TTRunFile = os.path.join(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\TTData-by-Intersection.xlsx')
    x1 = pd.ExcelFile(TTRunFile)
    x1.sheet_names
    Dat = x1.parse(TimePeriod, index_col=0,nrows= 5,usecols=['ClosestInt', 'DistFromSpd', 'DistBtwPnt', 'TimeDiff', 'SMS_mph',
       'SegName'])
    Dat = pd.merge(Dat,VissimDat,left_on=['ClosestInt'],right_on=['VEHICLETRAVELTIMEMEASUREMENT'], how ='left')
    Dat.loc[:,'DiffInTravelTime'] = Dat.TimeDiff- Dat.WeightedVissimTT
    Dat.rename(columns={'TimeDiff':'ObsTravelTime'},inplace=True)
    Dat = Dat[['TTSegNm','ObsTravelTime','WeightedVissimTT','DiffInTravelTime']] 
    Dat = Dat.round(1)

    return(Dat)

FinTTDat = {}
FinTTDat['ExistingAM_EB_TT']  =ProcessObsData('AM_EB',ExistingAMDat)
FinTTDat['ExistingAM_WB_TT'] = ProcessObsData('AM_WB',ExistingAMDat).sort_index(ascending=False)
FinTTDat['ExistingPM_EB_TT'] = ProcessObsData('PM_EB',ExistingPMDat)
FinTTDat['ExistingPM_WB_TT'] = ProcessObsData('PM_WB',ExistingPMDat).sort_index(ascending=False)


PathToKeyVal = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
OutFi = "Report-TT-GEH-Results.xlsx"
OutFi = os.path.join(PathToKeyVal,OutFi)

writer=pd.ExcelWriter(OutFi)
startrow1 = 1
for key,val in FinTTDat.items():
    val.to_excel(writer,'TTResults', startrow = startrow1+3)
    worksheet = writer.sheets['TTResults']
    worksheet.cell(startrow1+2, 1, key)
    startrow1 = worksheet.max_row
writer.save()


