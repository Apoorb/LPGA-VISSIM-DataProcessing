# -*- coding: utf-8 -*-
"""
Created on Fri Sep 20 08:22:23 2019

@author: abibeka
Purpose: Create Appendix for the report
"""
import glob
import pandas as pd    
import os
import re
    
ExistingFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Existing\*_Vehicle Travel Time Results.att')
NoBuildFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\No Build\*_Vehicle Travel Time Results.att')
DCDIFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCDI\*_Vehicle Travel Time Results.att')
DCMIFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCMI\*_Vehicle Travel Time Results.att')

file = ExistingFile[0]
def TTData_Subset(file, SegKeyValFile='TravelTimeKeyValuePairs.xlsx', IsBuid =False):
    '''
    file: TT result file
    SegKeyVal: Segment Name Key Value Pair
    '''
    PathToVis = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
    TimeConvFi = pd.ExcelFile(os.path.join(PathToVis,SegKeyValFile))
    TimeConvFi.sheet_names
    if(IsBuid):
        TTSegLaneDat = TimeConvFi.parse('Build')
    else:
        # Get The Travel time (TT) segment and # of lane data
        TTSegLaneDat = TimeConvFi.parse('Existing')

    TTSegLaneDat.fillna('',inplace=True)
    TTSegLaneDat.NumLanes = TTSegLaneDat.NumLanes.apply(lambda x: x if x == '' else '_{}_lanes'.format(int(x)))
    TTSegLaneDat.loc[:,'Segment Name'] = TTSegLaneDat.apply(lambda x: '{}: {}{}'.format(x[0],x[1],x[2]),axis=1)
    TTSegLaneDat = TTSegLaneDat[['SegNO','Segment Name']]
    
    
    Dat1=pd.read_csv(file,sep =';',skiprows=17)
    Dat1.columns
    Dat1.rename(columns={'$VEHICLETRAVELTIMEMEASUREMENTEVALUATION:SIMRUN': 'SimRun',
            'VEHICLETRAVELTIMEMEASUREMENT': 'SegNO',
            'TRAVTM(ALL)':'Travel Time','VEHS(ALL)':'Veh','DISTTRAV(ALL)': 'Distance'},inplace=True)
    mask=Dat1["SimRun"]=="AVG"
    Dat1 = Dat1[mask]
    Dat1 = Dat1.merge(TTSegLaneDat,left_on='SegNO',right_on='SegNO',how='left')
    Dat1 = Dat1[['SimRun','TIMEINT','Segment Name','Veh','Travel Time','Distance']]
    return Dat1


t =TTData_Subset(DCDIFile[0], SegKeyValFile='TravelTimeKeyValuePairs.xlsx', IsBuid =True)


DataDict = {
        'Existing AM': ExistingFile[0],
        'Existing PM': ExistingFile[1],
        'No-Build AM': NoBuildFile[0],
        'No-Build PM': NoBuildFile[1],
        'DCDI AM': DCDIFile[0],
        'DCDI PM': DCDIFile[1],
        'DCMI AM': DCMIFile[0],
        'DCMI PM': DCMIFile[1],
        }


OutFi = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\Appendix_Vissim.xlsx'



writer=pd.ExcelWriter(OutFi)

for name, file, in DataDict.items():
    startrow1 = 0
    tag = file.split('\\')[-2]
    match_build = re.search('[DCDI/DCMI]',tag)
    if(match_build):
        Dat = TTData_Subset(file,IsBuid =True)
    else:
        Dat = TTData_Subset(file,IsBuid =False)
    Dat.to_excel(writer,'TravelTimRes__{}'.format(name),startrow = startrow1+1,index=False)
    worksheet = writer.sheets['TravelTimRes__{}'.format(name)]
    worksheet.cell(startrow1+1, 1, name)
    startrow1 = worksheet.max_row





def ReadExistNodeRes(file):
    Dat=pd.read_csv(file,sep =';',skiprows=1,comment='*')
    Dat.columns
    #Use Avg values only
    Dat.rename(columns={'$MOVEMENTEVALUATION:SIMRUN':'Run','TIMEINT':'Time',
                        'MOVEMENT':'Movement','QLEN':'QLen','QLENMAX':'QLenMax','VEHS(ALL)':'Vehs',
                        'VEHDELAY(ALL)':'VehDelay'},inplace=True)
    Dat = Dat[(Dat['Run'] == 'AVG')]
    Dat = Dat[['Run','Time','Movement','QLen', 'QLenMax','Vehs','VehDelay']]
    return Dat


ExistingFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Existing\*Node Results.att')
NoBuildFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\No Build\*Node Results.att')
DCDCFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCDI\*Node Results.att')
DCMIFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCMI\*Node Results.att')

DataDict1 = {
        'Existing AM': ExistingFile[0],
        'Existing PM': ExistingFile[1],
        'No-Build AM': NoBuildFile[0],
        'No-Build PM': NoBuildFile[1],
        'DCDI AM': DCDCFile[0],
        'DCDI PM': DCDCFile[1],
        'DCMI AM': DCMIFile[0],
        'DCMI PM': DCMIFile[1],
        }

file = NoBuildFile[1]
for name, file, in DataDict1.items():
    startrow1 = 0
    Dat = ReadExistNodeRes(file)
    Dat.to_excel(writer,'NodeEvalRes_{}'.format(name),startrow = startrow1+1,index=False)
    worksheet = writer.sheets['NodeEvalRes_{}'.format(name)]
    worksheet.cell(startrow1+1, 1, name)
    startrow1 = worksheet.max_row
writer.save()    
