# -*- coding: utf-8 -*-
"""
Created on Mon Aug 19 13:47:01 2019

@author: abibeka
Read data from travel time files
"""

import os
import pandas as pd
import numpy as np
import glob
import datetime

PathToExist = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Existing'
ExistingPMfi = '20834_Existing_PM--C1C2C3C4C5C6C7_Vehicle Travel Time Results.att'
ExistingPMfi = os.path.join(PathToExist,ExistingPMfi)
ExistingAMfi ='20834_Existing_AM--C1C2aC3C4C5C6C7C8_Vehicle Travel Time Results.att'
ExistingAMfi = os.path.join(PathToExist,ExistingAMfi)



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
ExistingPMDat=pd.read_csv(ExistingPMfi,sep =';',skiprows=17)
ExistingPMDat.columns
ExistingPMDat.rename(columns={'TRAVTM(ALL)':'VissimTT','VEHS(ALL)':'Veh'},inplace=True)
mask=ExistingPMDat["$VEHICLETRAVELTIMEMEASUREMENTEVALUATION:SIMRUN"]=="AVG"
ExistingPMDat = ExistingPMDat[mask]
ExistingPMDat["TTSegNm"]=ExistingPMDat['VEHICLETRAVELTIMEMEASUREMENT'].apply(TTSegName)
mask = ExistingPMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100']) 
ExistingPM_EBDat = ExistingPMDat[mask]
mask2 = ExistingPMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100','8100-9000'])# 
ExistingPM_WBDat = ExistingPMDat[mask2]
ExistingPM_EBDat.loc[:,'VissimTotalTT'] = ExistingPM_EBDat.VissimTT * ExistingPM_EBDat.Veh
ExistingPM_WBDat.loc[:,'VissimTotalTT'] = ExistingPM_WBDat.VissimTT * ExistingPM_WBDat.Veh

ExistingPM_EBDat = ExistingPM_EBDat.groupby(['VEHICLETRAVELTIMEMEASUREMENT','TTSegNm'])['VissimTotalTT','Veh'].sum().reset_index()
ExistingPM_WBDat = ExistingPM_WBDat.groupby(['VEHICLETRAVELTIMEMEASUREMENT','TTSegNm'])['VissimTotalTT','Veh'].sum().reset_index()
ExistingPM_EBDat.loc[:,'WeightedVissimTT'] = ExistingPM_EBDat.VissimTotalTT/ExistingPM_EBDat.Veh
ExistingPM_WBDat.loc[:,'WeightedVissimTT'] = ExistingPM_WBDat.VissimTotalTT/ExistingPM_WBDat.Veh


ExistingAMDat=pd.read_csv(ExistingAMfi,sep =';',skiprows=17)
ExistingAMDat.columns
ExistingAMDat.rename(columns={'TRAVTM(ALL)':'VissimTT','VEHS(ALL)':'Veh'},inplace=True)
mask=ExistingAMDat["$VEHICLETRAVELTIMEMEASUREMENTEVALUATION:SIMRUN"]=="AVG"
ExistingAMDat = ExistingAMDat[mask]
ExistingAMDat["TTSegNm"]=ExistingAMDat['VEHICLETRAVELTIMEMEASUREMENT'].apply(TTSegName)
mask = ExistingAMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100']) 
ExistingAMDat = ExistingAMDat[mask]
ExistingAMDat.loc[:,'VissimTotalTT'] = ExistingAMDat.VissimTT * ExistingAMDat.Veh
ExistingAMDat = ExistingAMDat.groupby(['VEHICLETRAVELTIMEMEASUREMENT','TTSegNm'])['VissimTotalTT','Veh'].sum().reset_index()
ExistingAMDat.loc[:,'WeightedVissimTT'] = ExistingAMDat.VissimTotalTT/ExistingAMDat.Veh

#*********************************************************************************

#TTRun Files
#*********************************************************************************
TTRunFile = path.join(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\TTData-by-Intersection.xlsx')
x1 = pd.ExcelFile(TTRunFile)
x1.sheet_names
ExistPM_EB_Run = x1.parse('PM_EB',index_col=0,nrows= 5,usecols=['ClosestInt', 'DistFromSpd', 'DistBtwPnt', 'TimeDiff', 'SMS_mph',
       'SegName'])
ExistPM_WB_Run = x1.parse('PM_WB',index_col=0,nrows= 5,usecols=['ClosestInt', 'DistFromSpd', 'DistBtwPnt', 'TimeDiff', 'SMS_mph',
       'SegName'])
    
ExistAM_EB_Run = x1.parse('AM_EB',index_col=0,nrows= 5,usecols=['ClosestInt', 'DistFromSpd', 'DistBtwPnt', 'TimeDiff', 'SMS_mph',
       'SegName'])
ExistAM_WB_Run = x1.parse('AM_WB',index_col=0,nrows= 5,usecols=['ClosestInt', 'DistFromSpd', 'DistBtwPnt', 'TimeDiff', 'SMS_mph',
       'SegName'])
    
    

ExistPM_EB_Run = pd.merge(ExistPM_EB_Run,ExistingPM_EBDat, left_on=['ClosestInt'],right_on=['VEHICLETRAVELTIMEMEASUREMENT'], how ='left')
ExistPM_WB_Run = pd.merge(ExistPM_WB_Run,ExistingPM_EBDat, left_on=['ClosestInt'],right_on=['VEHICLETRAVELTIMEMEASUREMENT'], how ='left')
ExistPM_EB_Run.TimeDiff- ExistPM_EB_Run.WeightedVissimTT
ExistPM_WB_Run.TimeDiff- ExistPM_WB_Run.WeightedVissimTT

ExistAM_EB_Run = pd.merge(ExistAM_EB_Run,ExistingAMDat, left_on=['ClosestInt'],right_on=['VEHICLETRAVELTIMEMEASUREMENT'], how ='left')
ExistAM_WB_Run = pd.merge(ExistAM_WB_Run,ExistingAMDat, left_on=['ClosestInt'],right_on=['VEHICLETRAVELTIMEMEASUREMENT'], how ='left')
ExistAM_EB_Run.TimeDiff- ExistAM_EB_Run.WeightedVissimTT
ExistAM_WB_Run.TimeDiff- ExistAM_WB_Run.WeightedVissimTT

