# -*- coding: utf-8 -*-
"""
Created on Tue Sept 3 01:31:07 2019

@author: abibeka
# Get Density on I-95 (After Validation)
"""

#0.0 Housekeeping. Clear variable space
from IPython import get_ipython  #run magic commands
ipython = get_ipython()
ipython.magic("reset -f")
ipython = get_ipython()



import os
import pandas as pd
import numpy as np
import subprocess
import seaborn as sns
import matplotlib.pyplot as plt


# Get the time keys to convert Vissim intervals to Hours of the day
PathToVis = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
TimeConvFi  = 'TravelTimeKeyValuePairs.xlsx'
TimeConvFi = pd.ExcelFile(os.path.join(PathToVis,TimeConvFi))
TimeConvFi.sheet_names
TimeKeys = TimeConvFi.parse('TimeVissimKey')



# Define Sort Order
I95_Segs = [     'NB I-95 (US92 to NB OffRamp)',
                'NB I-95 (NB OffRamp to NB LoopRamp)',
                'NB I-95 ( NB LoopRamp to NB On-Ramp)',
                'NB I-95 (NB On-Ramp to SR40)',
                'SB I-95 (SR40 to SB OffRamp)',
                'SB I-95 (SB OffRamp to SB LoopRamp)',
                'SB I-95 (SB LoopRamp to SB On-Ramp)',
                'SB I-95 (SB On-Ramp to US92)'
               ]
TTSegLaneDat.loc[:,'SegName'] = pd.Categorical(TTSegLaneDat.SegName,I95_Segs)

TTSegLaneDat.sort_values('SegName',inplace=True)



# Get the Weighted Density by Segments
def PreProcessVissimDensity(file, SegKeyVal = TTSegLaneDat):
    '''
    file : VISSIM Results file
    SegKeyVal : Key value pair for segment # and TT segment name
    Summarize Vissim Travel time results
    '''
    ExistingAMDat=pd.read_csv(file,sep =';',skiprows=17)
    ExistingAMDat.columns
    ExistingAMDat.rename(columns={'TRAVTM(ALL)':'VissimTT','VEHS(ALL)':'Veh','DISTTRAV(ALL)':'Len'},inplace=True)
    mask=ExistingAMDat["$VEHICLETRAVELTIMEMEASUREMENTEVALUATION:SIMRUN"]=="AVG"
    ExistingAMDat = ExistingAMDat[mask]
    ExistingAMDat = SegKeyVal.merge(ExistingAMDat,left_on=['SegNO'],right_on = ['VEHICLETRAVELTIMEMEASUREMENT'],how= 'left')
    ExistingAMDat.TIMEINT = pd.Categorical(ExistingAMDat.TIMEINT,['900-1800','1800-2700','2700-3600','3600-4500',
                                        '4500-5400','5400-6300','6300-7200','7200-8100',
                                        '8100-9000','9000-9900','9900-10800','10800-11700'])
    ExistingAMDat.loc[:,"VissimSMS"] = (ExistingAMDat['Len']/ExistingAMDat['VissimTT']/1.47).round(1)
    #Get flow rate and density
    ExistingAMDat.loc[:,'FlowRate'] = ExistingAMDat.Veh *4
    ExistingAMDat.loc[:,'DensityPerLane'] = (ExistingAMDat.FlowRate/ ExistingAMDat.VissimSMS/ExistingAMDat.NumLanes).round(1)
    ExistingAMDat.loc[:,'LenByDensity'] = ExistingAMDat.DensityPerLane *ExistingAMDat.Len
    ExistingAMDat.columns
    DensityData = ExistingAMDat.groupby(['TIMEINT','SegName'])['Len','LenByDensity'].sum().reset_index()
    DensityData.loc[:,'WeightedDensity'] = (DensityData.LenByDensity/ DensityData.Len).round(1)
    return(DensityData)


# VISSIM File
#*********************************************************************************
PathToExist = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Existing'
ExistingPMfi = '20834_Existing_PM--C1C2C3C4C5C6C7_Vehicle Travel Time Results.att'
ExistingPMfi = os.path.join(PathToExist,ExistingPMfi)
ExistingAMfi ='20834_Existing_AM--C1C2aC3C4C5C6C7C8_Vehicle Travel Time Results.att'
ExistingAMfi = os.path.join(PathToExist,ExistingAMfi)

file = ExistingAMfi
SegKeyVal = TTSegLaneDat
DenAM = PreProcessVissimTT(ExistingAMfi,SegKeyVal)
DenPM = PreProcessVissimTT(ExistingPMfi,SegKeyVal)







