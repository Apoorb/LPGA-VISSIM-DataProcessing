# -*- coding: utf-8 -*-
"""
Created on Tue Dec 10 10:00:03 2019

@author: abibeka
Calibration Table for 12th Street
"""


#0.0 Housekeeping. Clear variable space
from IPython import get_ipython  #run magic commands
ipython = get_ipython()
ipython.magic("reset -f")
ipython = get_ipython()



import os
import pandas as pd

os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\12th-Street-TransitWay\Results')
#os.chdir(r'H:\20\20548 - Arlington County Engineering On-Call\009 - 12th Street Transitway Extension\vissim\Results')



x1 = pd.ExcelFile(r'ResultsNameMap.xlsx')
x1.sheet_names
TTMap = x1.parse('TTMap')

Field_TT = x1.parse('TTFieldCalibMap')

#******************************************************************************************************************************************************************
def PreProcessVissimTT(file,TTMap = TTMap,Peak= "AM Peak", SimRun = "AVG"):
    #Ignore the variable names. Don't actually mean ExistingAMDat
    # Define PM file. Would later use to figure out if a file for PM 
    # This is hard coding. Would break if the name of PM file is changed.
    # Read VISSIM results
    ExistingAMDat=pd.read_csv(file,sep =';',comment = "*",skiprows=1)
    ExistingAMDat.columns
    ExistingAMDat.rename(columns={'TRAVTM(ALL)':'VissimTT','VEHS(ALL)':'Veh'},inplace=True)
    mask=ExistingAMDat["$VEHICLETRAVELTIMEMEASUREMENTEVALUATION:SIMRUN"]==SimRun
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
    Dat.loc[:,"Peak"] = Peak
    assert((Peak=="AM Peak") | (Peak=="PM Peak" ))
    return(Dat)


    
#*********************************************************************************
# Specify Files
#*********************************************************************************

TTFile_AM = r'./RawVissimOutput/20548_2019_am-existing_V12_Vehicle Travel Time Results.att'
file = TTFile_AM
TTFile_PM =  r'./RawVissimOutput/20548_2019_pm-existing_V6_Vehicle Travel Time Results.att'


#*********************************************************************************
# Call Function
#*********************************************************************************
TT_Existing_AM = PreProcessVissimTT(file = TTFile_AM,TTMap = TTMap, Peak = "AM Peak" ,SimRun= "AVG")
TT_Existing_PM = PreProcessVissimTT(file = TTFile_PM,TTMap = TTMap, Peak = "PM Peak", SimRun= "AVG")
TT_Existing = pd.concat([TT_Existing_AM,TT_Existing_PM])
merge_on =["TT_No","Peak"]
Field_TT.rename(columns = {'Vissim_TT_No':'TT_No'}, inplace=True)
Calibration_Dat = Field_TT.merge(TT_Existing,left_on = merge_on, right_on = merge_on, how ='left')
Calibration_Dat.rename(columns = {'WeightedVissimTT':'VissimAvgTT'}, inplace=True)
Calibration_Dat = Calibration_Dat[['Peak','Direction','SegmentName','ObsAvgTT','VissimAvgTT']]
Calibration_Dat.loc[:,'TT_Difference'] = Calibration_Dat.ObsAvgTT - Calibration_Dat.VissimAvgTT

Calibration_Dat.set_index(['Peak','Direction','SegmentName'],inplace=True)

OutFi = r'TT_DataCol_Calibration.xlsx'
writer=pd.ExcelWriter(OutFi)
Calibration_Dat.to_excel(writer, "TT_Calibration" ,na_rep='-')
writer.save() 
#TT_Existing_PM = PreProcessVissimTT(file = TTFile_PM,TTMap = TTMap)

#TT_Dat_Existing = TT_Existing_AM.merge(TT_Existing_PM,left_on=['TT_No','TT_Name','IsTransit'], right_on=['TT_No','TT_Name','IsTransit'], how ='inner', suffixes = ('AM','PM'))

