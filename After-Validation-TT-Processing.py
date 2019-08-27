# -*- coding: utf-8 -*-
"""
Created on Tue Aug 27 08:04:07 2019

@author: abibeka
# Get Travel time on LPGA and I-95 (After Validation)
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

# VISSIM File
#*********************************************************************************
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
    Nm = TTSeg[x]
    return Nm




def PreProcessVissimTT(file, SegIN):
    '''
    file : VISSIM Results file
    SegIN : Specify segments for which results are needed
    '''
    ExistingAMDat=pd.read_csv(file,sep =';',skiprows=17)
    ExistingAMDat.columns
    ExistingAMDat.rename(columns={'TRAVTM(ALL)':'VissimTT','VEHS(ALL)':'Veh'},inplace=True)
    mask=ExistingAMDat["$VEHICLETRAVELTIMEMEASUREMENTEVALUATION:SIMRUN"]=="AVG"
    ExistingAMDat = ExistingAMDat[mask]
    ExistingAMDat["TTSegNm"]=ExistingAMDat['VEHICLETRAVELTIMEMEASUREMENT'].apply(TTSegName)
    ExistingAMDat.TIMEINT = pd.Categorical(ExistingAMDat.TIMEINT,['900-1800','1800-2700','2700-3600','3600-4500',
                                        '4500-5400','5400-6300','6300-7200','7200-8100',
                                        '8100-9000','9000-9900','9900-10800','10800-11700'])
    
    mask = ExistingAMDat.TTSegNm.isin(I95_Segs)
    ExistingAMDat = ExistingAMDat[mask]
    ExistingAMDat.loc[:,"VissimSMS"] = (ExistingAMDat['DISTTRAV(ALL)']/ExistingAMDat['VissimTT']/1.47).round(1)
    mask1 = ~ExistingAMDat['VEHICLETRAVELTIMEMEASUREMENT'].isin([13,14])
    TTData = ExistingAMDat[mask1]
    return(TTData)

I95_Segs = [     'NB I-95 (US92 to NB OffRamp)',
                'NB I-95 (NB OffRamp to NB LoopRamp)',
                'NB I-95 ( NB LoopRamp to NB On-Ramp)',
                'NB I-95 (NB On-Ramp to SR40)',
                'SB I-95 (SR40 to SB OffRamp)',
                'SB I-95 (SB OffRamp to SB LoopRamp)',
                'SB I-95 (SB LoopRamp to SB On-Ramp)',
                'SB I-95 (SB On-Ramp to US92)'
               ]
  
LPGASeg=['EB LPGA (Tomoka Rd to I-95 SB Ramp)',
    'EB LPGA (I-95 SB Ramp to I-95 NB Ramp)',
    'EB LPGA (I-95 NB Ramp to Technology Blvd)', 
    'EB LPGA (Technology Blvd to Willamson Blvd)',
    'EB LPGA (Willamson Blvd to Clyde-Morris Blvd)',
    'WB LPGA (Clyde-Morris Blvd to Willamson Blvd)', 
    'WB LPGA (Willamson Blvd to Technology Blvd)', 
    'WB LPGA (Technology Blvd to I-95 NB Ramp)',
    'WB LPGA (I-95 NB Ramp to I-95 SB Ramp)',
    'WB LPGA (I-95 SB Ramp to Tomoka Rd)']
SegIn1 = I95_Segs+ LPGASeg
file = ExistingAMfi
TTAM = PreProcessVissimTT(ExistingAMfi,SegIn1)
TTPM = PreProcessVissimTT(ExistingPMfi,SegIn1)

TTAM.loc[:,'AnalysisPeriod'] = 'AM'
TTPM.loc[:,'AnalysisPeriod'] = 'PM'

Vissim_TT = pd.concat([TTAM,TTPM],sort=False)
Vissim_TT.columns
Vissim_TT = Vissim_TT[['AnalysisPeriod','TIMEINT','TTSegNm','VissimSMS','Veh']]

Vissim_TT.TTSegNm = pd.Categorical(Vissim_TT.TTSegNm,SegIn1)
    



