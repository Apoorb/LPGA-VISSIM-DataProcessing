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
# Use consistent naming in VISSIM
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
            12:'WB LPGA (I-95 SB Ramp to Tomoka Rd)',
            13:'SB I-95',
            14:'NB I-95',
            15:'SB I-95 (SR40 to SB OffRamp)',
            16:'SB I-95 (SB OffRamp to SB LoopRamp)',
            17:'SB I-95 (SB LoopRamp to SB On-Ramp)',
            18:'SB I-95 (SB On-Ramp to US92)',
            19:'NB I-95 (US92 to NB OffRamp)',
            20:'NB I-95 (NB OffRamp to NB LoopRamp)',
            21:'NB I-95 ( NB LoopRamp to NB On-Ramp)',
            22:'NB I-95 (NB On-Ramp to SR40)'}
    if(x<23):
        Nm = TTSeg[x]
    else: Nm = None
    return Nm

# VISSIM File
#*********************************************************************************
PathToExist = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\July-6-2020\Models---July-2020\VISSIM\Existing'
ExistingPMfi = '20834_Existing_PM_Vehicle Travel Time Results.att'
ExistingPMfi = os.path.join(PathToExist,ExistingPMfi)
ExistingAMfi ='20834_Existing_AM_Vehicle Travel Time Results.att'
ExistingAMfi = os.path.join(PathToExist,ExistingAMfi)

def PreProcessVissimTT(file = ExistingAMfi):
    # Define PM file. Would later use to figure out if a file for PM 
    # This is hard coding. Would break if the name of PM file is changed.
    RefFile = ExistingPMfi
    # Read VISSIM results
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
    #'4500-5400','5400-6300','6300-7200','7200-8100' are peak periods
    if (file ==ExistingPMfi):
        mask1 = (ExistingAMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100'])) & (~ExistingAMDat.TTSegNm.isin(WB_TTSegs))
        #Can include '8100-9000' as WB TT Run includes 15 min after the peak
        mask2 = (ExistingAMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100','8100-9000'])) & (ExistingAMDat.TTSegNm.isin(WB_TTSegs))
        mask = mask1 | mask2
    else:
        mask = ExistingAMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100']) 
    ExistingAMDat = ExistingAMDat[mask]
    # Get weighted average over the 4/5 intervals
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
    #Read the observed TT data
    TTRunFile = os.path.join(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\TTData-by-Intersection.xlsx')
    x1 = pd.ExcelFile(TTRunFile)
    x1.sheet_names
    Dat = x1.parse(TimePeriod, index_col=0,nrows= 5,usecols=['ClosestInt', 'DistFromSpd', 'DistBtwPnt', 'TimeDiff', 'SMS_mph',
       'SegName'])
    #Merge with ViSSIM TT data
    Dat = pd.merge(Dat,VissimDat,left_on=['ClosestInt'],right_on=['VEHICLETRAVELTIMEMEASUREMENT'], how ='outer')
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


#*********************************************************************************
# Write to excel
#*********************************************************************************

PathToKeyVal = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\July-6-2020\Models---July-2020\VISSIM'
OutFi = "Report-TT-GEH-Results.xlsx"
OutFi = os.path.join(PathToKeyVal,OutFi)

writer=pd.ExcelWriter(OutFi)
startrow1 = 1
for key,val in FinTTDat.items():
    # Write tables on same sheet wih 2 row spacing
    val.to_excel(writer,'TTResults', startrow = startrow1+3)
    worksheet = writer.sheets['TTResults']
    worksheet.cell(startrow1+2, 1, key)
    startrow1 = worksheet.max_row
writer.save()


