# -*- coding: utf-8 -*-
"""
Created on Mon Aug 19 13:47:01 2019

@author: abibeka
Data collection point processing for calibration
"""

#0.0 Housekeeping. Clear variable space
from IPython import get_ipython  #run magic commands
ipython = get_ipython()
ipython.magic("reset -f")
ipython = get_ipython()



import os
import pandas as pd

# Use consistent naming in VISSIM
#os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\12th-Street-TransitWay\Results')
os.chdir(r'H:\20\20548 - Arlington County Engineering On-Call\009 - 12th Street Transitway Extension\vissim\Results')


x1 = pd.ExcelFile(r'ResultsNameMap.xlsx')
x1.sheet_names
DatColMap = x1.parse('DatColMap')
FieldDataColMap = x1.parse('FieldDataColMap')



    
def PreProcessVissimDataCol(file,DatColMap = DatColMap,Peak= "AM Peak"):
    #Ignore the variable names. Don't actually mean ExistingAMDat
    # Define PM file. Would later use to figure out if a file for PM 
    # This is hard coding. Would break if the name of PM file is changed.
    # Read VISSIM results
    ExistingAMDat=pd.read_csv(file,sep =';',comment = "*",skiprows=1)
    ExistingAMDat.columns
    ExistingAMDat.rename(columns={'SPEEDAVGARITH(ALL)':'VissimAvgSpd'},inplace=True)
    mask=ExistingAMDat["$DATACOLLECTIONMEASUREMENTEVALUATION:SIMRUN"]=="AVG"
    ExistingAMDat = ExistingAMDat[mask]
    ExistingAMDat = ExistingAMDat.merge(DatColMap,left_on = 'DATACOLLECTIONMEASUREMENT', right_on = 'DataColNo', how ='left')
    mask1 = ExistingAMDat.TIMEINT.isin(['900-4500']) 
    ExistingAMDat = ExistingAMDat[mask1]
    ExistingAMDat = ExistingAMDat[['DataColNo','DatColNm','VissimAvgSpd']]
    # Get weighted average over the 4/5 intervals
    ExistingAMDat = ExistingAMDat.round(1)
    ExistingAMDat.loc[:,"Peak"] = Peak
    assert((Peak=="AM Peak") | (Peak=="PM Peak" ))
    return(ExistingAMDat)
    

#*********************************************************************************
# Specify Files
#*********************************************************************************
x1 = pd.ExcelFile('CollectedData/ATR_AM_PM_Speeds.xlsx')
x1.sheet_names
Field_DatCol = x1.parse('SAvgSpdField',index_col =0)
Field_DatCol =Field_DatCol.merge(FieldDataColMap,left_on =["Location","Dir"],right_on = ["Segment","Direction"]).drop(columns =["Segment","Dir"])
Field_DatCol.rename(columns = {'AvgSpd':'ObsAvgSpd'},inplace=True)

#*********************************************************************************
# Specify Files
#*********************************************************************************
DatColFile_AM = r'./RawVissimOutput/20548_2019_am-existing_V6_Data Collection Results.att'
file  = DatColFile_AM
DatColFile_PM = DatColFile_AM
DatCol_Existing_AM = PreProcessVissimDataCol(file = DatColFile_PM,DatColMap = DatColMap,Peak= "AM Peak")
DatCol_Existing_PM = PreProcessVissimDataCol(file = DatColFile_PM,DatColMap = DatColMap,Peak= "PM Peak")    

#*********************************************************************************
# Call Function
#*********************************************************************************
DatCol_Existing = pd.concat([DatCol_Existing_AM,DatCol_Existing_PM])
merge_on =["DataColNo","Peak"]
Calibration_Dat = Field_DatCol.merge(DatCol_Existing,left_on = merge_on, right_on = merge_on, how ='left')
Calibration_Dat = Calibration_Dat[['Peak','Direction',"DatColNm",'DataColNo','ObsAvgSpd','VissimAvgSpd']]
Calibration_Dat.sort_values(['Peak','DataColNo'],inplace=True)
Calibration_Dat.loc[:,'Spd_Difference'] = Calibration_Dat.ObsAvgSpd - Calibration_Dat.VissimAvgSpd
Calibration_Dat.set_index(['Peak','Direction','DatColNm'],inplace=True)
Calibration_Dat = Calibration_Dat.round(1)
OutFi = r'TT_DataCol_Calibration.xlsx'
writer=pd.ExcelWriter(OutFi, mode= 'a')
Calibration_Dat.to_excel(writer, "Speed_Calibration" ,na_rep='-')
writer.save() 