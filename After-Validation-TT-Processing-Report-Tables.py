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
import re
# Use consistent naming in VISSIM

def TTSegName(x):
    TTSeg= {
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
    if ((x>2) & (x < 13)):
        Nm = TTSeg[x]
    else:
        Nm = None
    return Nm

# VISSIM File
#*********************************************************************************
NoBuildFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\No Build\*_Vehicle Travel Time Results.att')
DCDIFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCDI\*_Vehicle Travel Time Results.att')
DCMIFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCMI\*_Vehicle Travel Time Results.att')


def PreProcessVissimTT(file):
    #Ignore the variable names. Don't actually mean ExistingAMDat
    # Define PM file. Would later use to figure out if a file for PM 
    # This is hard coding. Would break if the name of PM file is changed.
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
    match_AM = re.search('_AM',file)
    if (match_AM):
        mask = ExistingAMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100']) 
    else:
        mask1 = (ExistingAMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100'])) & (~ExistingAMDat.TTSegNm.isin(WB_TTSegs))
        #Can include '8100-9000' as WB TT Run includes 15 min after the peak
        mask2 = (ExistingAMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100','8100-9000'])) & (ExistingAMDat.TTSegNm.isin(WB_TTSegs))
        mask = mask1 | mask2
    ExistingAMDat = ExistingAMDat[mask]
    # Get weighted average over the 4/5 intervals
    ExistingAMDat.loc[:,'VissimTotalTT'] = ExistingAMDat.VissimTT * ExistingAMDat.Veh
    ExistingAMDat = ExistingAMDat.groupby(['VEHICLETRAVELTIMEMEASUREMENT','TTSegNm'])['VissimTotalTT','Veh'].sum().reset_index()
    ExistingAMDat.loc[:,'WeightedVissimTT'] = ExistingAMDat.VissimTotalTT/ExistingAMDat.Veh
    Dat = ExistingAMDat[['TTSegNm','WeightedVissimTT']] 
    Dat = Dat.round(1)
    return(Dat)
    
NoBuild_AM = PreProcessVissimTT(file = NoBuildFile[0])
NoBuild_PM = PreProcessVissimTT(file = NoBuildFile[1])
Dat_NoBuild = NoBuild_AM.merge(NoBuild_PM,left_on='TTSegNm', right_on='TTSegNm', how ='inner', suffixes = ('AM','PM'))

DCDI_AM = PreProcessVissimTT(DCDIFile[0])
DCDI_PM = PreProcessVissimTT(DCDIFile[1])
Dat_DCDI = DCDI_AM.merge(DCDI_PM,left_on='TTSegNm', right_on='TTSegNm', how ='inner', suffixes = ('AM','PM'))

DCMI_AM = PreProcessVissimTT(DCMIFile[0])
DCMI_PM = PreProcessVissimTT(DCMIFile[1])
Dat_DCMI = DCMI_AM.merge(DCMI_PM,left_on='TTSegNm', right_on='TTSegNm', how ='inner', suffixes = ('AM','PM'))



#*********************************************************************************
# Write to excel
#*********************************************************************************

PathToKeyVal = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
OutFi = "Report-TT-NoBuild-and-Build-Scenarios.xlsx"
OutFi = os.path.join(PathToKeyVal,OutFi)

writer=pd.ExcelWriter(OutFi)
startrow1 = 1
listDat = [Dat_NoBuild,Dat_DCDI,Dat_DCMI]
Names = ['No-Build','DCDI','DCMI']
for key,val in zip( Names,listDat):
    # Write tables on same sheet wih 2 row spacing
    val.to_excel(writer,'Build-NoBuild-TTRes', startrow = startrow1+3)
    worksheet = writer.sheets['Build-NoBuild-TTRes']
    worksheet.cell(startrow1+2, 1, key)
    startrow1 = worksheet.max_row
writer.save()


