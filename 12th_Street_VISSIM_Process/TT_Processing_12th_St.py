# -*- coding: utf-8 -*-
"""
Created on Mon Aug 19 13:47:01 2019

@author: abibeka
Read data from travel time and data collection files
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

#os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\12th-Street-TransitWay\Results')
os.chdir(r'H:\20\20548 - Arlington County Engineering On-Call\009 - 12th Street Transitway Extension\vissim\Results')



x1 = pd.ExcelFile(r'ResultsNameMap.xlsx')
x1.sheet_names
TTMap = x1.parse('TTMap')
DatColMap = x1.parse('DatColMap')

#******************************************************************************************************************************************************************
def PreProcessVissimTT(file,TTMap = TTMap):
    #Ignore the variable names. Don't actually mean ExistingAMDat
    # Define PM file. Would later use to figure out if a file for PM 
    # This is hard coding. Would break if the name of PM file is changed.
    # Read VISSIM results
    ExistingAMDat=pd.read_csv(file,sep =';',comment = "*",skiprows=1)
    ExistingAMDat.columns
    ExistingAMDat.rename(columns={'TRAVTM(ALL)':'VissimTT','VEHS(ALL)':'Veh'},inplace=True)
    mask=ExistingAMDat["$VEHICLETRAVELTIMEMEASUREMENTEVALUATION:SIMRUN"]=="AVG"
    ExistingAMDat = ExistingAMDat[mask]
    ExistingAMDat = ExistingAMDat.merge(TTMap,left_on = 'VEHICLETRAVELTIMEMEASUREMENT', right_on = 'TT_No', how ='left')
    mask1 = ExistingAMDat.TIMEINT.isin(['900-4500']) 
    ExistingAMDat = ExistingAMDat[mask1]
    ExistingAMDat = ExistingAMDat[['TT_No','TT_Name','IsTransit','Veh','VissimTT','DISTTRAV(ALL)']]
    # Get weighted average over the 4/5 intervals
    ExistingAMDat.loc[:,'VissimTotalTT'] = ExistingAMDat.VissimTT * ExistingAMDat.Veh
    ExistingAMDat = ExistingAMDat.groupby(['TT_No','TT_Name','IsTransit'])['VissimTotalTT','Veh'].sum().reset_index()
    ExistingAMDat.loc[:,'WeightedVissimTT'] = ExistingAMDat.VissimTotalTT/ExistingAMDat.Veh
    Dat = ExistingAMDat[['TT_No','TT_Name','IsTransit','WeightedVissimTT']] 
    Dat = Dat.round(1)
    return(Dat)
    
    
    
#******************************************************************************************************************************************************************
def PreProcessVissimDataCol(file,DatColMap = DatColMap):
    #Ignore the variable names. Don't actually mean ExistingAMDat
    # Define PM file. Would later use to figure out if a file for PM 
    # This is hard coding. Would break if the name of PM file is changed.
    # Read VISSIM results
    ExistingAMDat=pd.read_csv(file,sep =';',comment = "*",skiprows=1)
    ExistingAMDat.columns
    ExistingAMDat.rename(columns={'SPEEDAVGARITH(ALL)':'AvgSpeedArit'},inplace=True)
    mask=ExistingAMDat["$DATACOLLECTIONMEASUREMENTEVALUATION:SIMRUN"]=="AVG"
    ExistingAMDat = ExistingAMDat[mask]
    ExistingAMDat = ExistingAMDat.merge(DatColMap,left_on = 'DATACOLLECTIONMEASUREMENT', right_on = 'DataColNo', how ='left')
    mask1 = ExistingAMDat.TIMEINT.isin(['900-4500']) 
    ExistingAMDat = ExistingAMDat[mask1]
    ExistingAMDat = ExistingAMDat[['DataColNo','DatColNm','AvgSpeedArit']]
    # Get weighted average over the 4/5 intervals
    ExistingAMDat = ExistingAMDat.round(1)
    return(ExistingAMDat)
    
    
#*********************************************************************************
# Specify Files
#*********************************************************************************

TTFile_AM = r'./RawVissimOutput/20548_2019_am-existing_V5_Vehicle Travel Time Results.att'
#file = TTFile
DatColFile_AM = r'./RawVissimOutput/20548_2019_am-existing_V5_Data Collection Results.att'
file  = DatColFile_AM

#*********************************************************************************
# Call Function
#*********************************************************************************
TTFile_PM = TTFile_AM # change later when you get results
TT_Existing_AM = PreProcessVissimTT(file = TTFile_AM,TTMap = TTMap)
TT_Existing_PM = PreProcessVissimTT(file = TTFile_PM,TTMap = TTMap)
TT_Dat_Existing = TT_Existing_AM.merge(TT_Existing_PM,left_on=['TT_No','TT_Name','IsTransit'], right_on=['TT_No','TT_Name','IsTransit'], how ='inner', suffixes = ('AM','PM'))

DatColFile_PM = DatColFile_AM # change later when you get results

TTFile_PM = TTFile_AM # change later when you get results
DatCol_Existing_AM = PreProcessVissimDataCol(file = DatColFile_PM,DatColMap = DatColMap)
DatCol_Existing_PM = PreProcessVissimDataCol(file = DatColFile_PM,DatColMap = DatColMap)
DatCol_Dat_Existing = DatCol_Existing_AM.merge(DatCol_Existing_PM,left_on=['DataColNo','DatColNm'], right_on=['DataColNo','DatColNm'], how ='inner', suffixes = ('AM','PM'))

#*********************************************************************************
# Write to excel
#*********************************************************************************
OutFi = "Report-TT_DatCol-Existing-and-Build-Scenarios.xlsx"
OutFi = os.path.join(OutFi)

writer=pd.ExcelWriter(OutFi)
startrow1 = 1
listDat = [TT_Dat_Existing]
Names = ['Existing']
for key,val in zip( Names,listDat):
    # Write tables on same sheet wih 2 row spacing
    val.to_excel(writer,'12th_St-TTRes', startrow = startrow1+3)
    worksheet = writer.sheets['12th_St-TTRes']
    worksheet.cell(startrow1+2, 1, key)
    startrow1 = worksheet.max_row
    
startrow1 = 1
listDat = [DatCol_Dat_Existing]
Names = ['Existing']
for key,val in zip( Names,listDat):
    # Write tables on same sheet wih 2 row spacing
    val.to_excel(writer,'12th_St-DatColRes', startrow = startrow1+3)
    worksheet = writer.sheets['12th_St-DatColRes']
    worksheet.cell(startrow1+2, 1, key)
    startrow1 = worksheet.max_row
    
writer.save()

