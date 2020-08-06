# -*- coding: utf-8 -*-
"""
Created on Fri Sep  6 13:01:32 2019

@author: abibeka
Purpose: Compare VISSIM Volumes with collected volumes. Calc GEH
"""

#0.0 Housekeeping. Clear variable space
from IPython import get_ipython  #run magic commands
ipython = get_ipython()
ipython.magic("reset -f")
ipython = get_ipython()

#1 Import Required Packages
#*********************************************************************************
import os
import pandas as pd
import numpy as np
import subprocess 

os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\12th-Street-TransitWay\Results')
#os.chdir(r'H:\20\20548 - Arlington County Engineering On-Call\009 - 12th Street Transitway Extension\vissim\Results')
os.getcwd()

# Read the Mapping Data
#*********************************************************************************
x1 = pd.ExcelFile('ResultsNameMap.xlsx')
x1.sheet_names
NodeMap = x1.parse('NodeMap')  
DirMap = x1.parse('DirMap')


# Read VISSIM File
#*********************************************************************************
def ProcessVissimVolumes(file, NodeMap = NodeMap, DirMap = DirMap, SimRun = "AVG"):
    '''
    file: Node Evaluation file
    Returns: Cleaned resuls
    Read the Node-Evaluation File
    Ignore the ExistingPMfi. Can be used all AM and PM existing Node Evaluation
    '''
    ExistingPMDat=pd.read_csv(file,sep =';',skiprows=1,comment="*")
    ExistingPMDat.columns
    #Use Avg values only
    ExistingPMDat = ExistingPMDat[(ExistingPMDat['$MOVEMENTEVALUATION:SIMRUN'] ==SimRun)]
    #Only use motorized links
    Mask_Not4or5 = ~((ExistingPMDat['MOVEMENT\TOLINK\LINKBEHAVTYPE']==4) | (ExistingPMDat['MOVEMENT\TOLINK\LINKBEHAVTYPE']==5))
    ExistingPMDat = ExistingPMDat[Mask_Not4or5]
    ExistingPMDat.rename(columns={'MOVEMENT':'Mvmt'},inplace=True)
    ExPMMvMDat = ExistingPMDat.copy() #Something from previous code. Not needed here
    ExPMMvMDat.loc[:,"HourInt"] = 'Nan'
    ExPMMvMDat.loc[:,'HourInt'] = ExPMMvMDat.TIMEINT
    ExPMMvMDat.HourInt = pd.Categorical(ExPMMvMDat.HourInt,['900-4500']) # Something from previous code. Not needed here
    ExPMMvMDat.rename(columns= {'VEHS(ALL)':'VissimVol','MOVEMENT\DIRECTION':'VissimDir'},inplace=True)
    new = ExPMMvMDat.Mvmt.str.split(':',expand=True)
    ExPMMvMDat['Intersection'] = (new[0].str.split('-',n=1,expand=True)[0]).astype('int')
    ExPMMvMDat = ExPMMvMDat[['Intersection','VissimDir','VissimVol']]
    #Handle duplicate rows 
    ExPMMvMDat = ExPMMvMDat.groupby(['Intersection','VissimDir']).agg({'VissimVol':'sum'}).reset_index()
    ExPMMvMDat = ExPMMvMDat.merge(DirMap, left_on = 'VissimDir',right_on = 'VissimDir',how = 'left') 
    #--------------------------------------------------------------------------------------------------
    ExPMMvMDat = pd.merge(ExPMMvMDat,NodeMap,left_on=['Intersection'],
                 right_on=["NodeNum"], how = 'left')
    ExPMMvMDat.CardinalDir = pd.Categorical(ExPMMvMDat.CardinalDir,[
        'EBL','EBT','EBR','WBL','WBT','WBR','NBL','NBT','NBR','SBL','SBT','SBR'])
    ExPMMvMDat = ExPMMvMDat[['Intersection',"NodeName",'CardinalDir','VissimVol']]
    ExPMMvMDat.columns
    ExPMMvMDat =ExPMMvMDat[~ExPMMvMDat.CardinalDir.isna()]
    return(ExPMMvMDat)
 
# Read Observed TMC File
#*********************************************************************************   
def ProcessTMC_Data(file):
    '''
    file = Observed TMC volumes 
    Reshape data so if can be merged with VISSIM data
    '''
    TMC_dat= pd.read_csv(file,skiprows=23)
    TMC_dat = TMC_dat[TMC_dat.RECORDNAME=="Volume"]
    TMC_dat =TMC_dat[['INTID','EBL','EBT','EBR','WBL','WBT','WBR','NBL','NBT','NBR','SBL','SBT','SBR']]
    TMC_dat.set_index('INTID',inplace=True)
    TMC_dat = TMC_dat.stack().reset_index()
    TMC_dat.columns = ['Intersection','CardinalDir','ObsVol']
    return(TMC_dat)
   
# Merge VISSIM and Observed TMC Files
#*********************************************************************************
def MergeVISSIM_ObsTMC(VissimDat, ObsDat):
    '''
    Parameters
    ----------
    VissimDat : pd.DataFrame
        VISSIM TMC data.
    ObsDat : pd.DataFrame
        Observed TMC data.

    Returns 
    -------
    MergeDat : Merged VISSIM and Observed data

    '''
    mergeCol = ['Intersection','CardinalDir']
    MergeDat = pd.merge(VissimDat,ObsDat, left_on = mergeCol, right_on = mergeCol, how = 'left')
    MergeDat.CardinalDir = pd.Categorical(MergeDat.CardinalDir,[
        'EBL','EBT','EBR','WBL','WBT','WBR','NBL','NBT','NBR','SBL','SBT','SBR'])
    MergeDat = MergeDat.sort_values(['Intersection','CardinalDir'],axis=0)
    MergeDat.dtypes
    MergeDat.ObsVol = MergeDat.ObsVol.astype(int)
    MergeDat.loc[:,'GEH'] = np.sqrt(((MergeDat.VissimVol - MergeDat.ObsVol)**2)/ ((MergeDat.VissimVol + MergeDat.ObsVol)/2))
    MergeDat.loc[:,:'ObsVol'] = MergeDat.loc[:,:'ObsVol'].round()
    MergeDat.loc[:,:'VissimVol'] = MergeDat.loc[:,:'VissimVol'].round()
    MergeDat.loc[:,:'GEH'] = MergeDat.loc[:,:'GEH'].round(2)
    
    # Debug = MergeDat
    # Debug.loc[:,'Num'] = ((MergeDat.VissimVol - MergeDat.ObsVol)**2)
    # Debug.loc[:,'Denom'] = ((MergeDat.VissimVol + MergeDat.ObsVol)/2)
    # Debug.loc[:,'Ratio'] = MergeDat.loc[:,'Num']/  MergeDat.loc[:,'Denom']
    # Debug.loc[:,'Ratio_Sqrt'] = np.sqrt(MergeDat.loc[:,'Ratio'])
    return(MergeDat)

# Get Final GEH Values for AM and PM 
#*********************************************************************************
def FinalTable_GEH(AM_Res, PM_Res):
    #Get Final Table
    AM_Res = AM_Res[~pd.isna(AM_Res.GEH)]
    GEHBelow2= sum((AM_Res.loc[:,'GEH']<=2).values.ravel())/AM_Res.shape[0]
    GEHBelow5= sum((AM_Res.loc[:,'GEH']<=5).values.ravel())/AM_Res.shape[0]
    GEHBelow10= sum((AM_Res.loc[:,'GEH']<=10).values.ravel())/AM_Res.shape[0]
    PM_Res = PM_Res[~pd.isna(PM_Res.GEH)]
    GEHBelow2_pm= sum((PM_Res.loc[:,'GEH']<=2).values.ravel())/PM_Res.shape[0]
    GEHBelow5_pm= sum((PM_Res.loc[:,'GEH']<=5).values.ravel())/PM_Res.shape[0]
    GEHBelow10_pm= sum((PM_Res.loc[:,'GEH']<=10).values.ravel())/PM_Res.shape[0]
    
    FinTab = {'Time' : ['AM Peak Hour','AM Peak Hour','AM Peak Hour','PM Peak Hour','PM Peak Hour','PM Peak Hour'],
    'Calibration Metric' : ['GEH Below 2','GEH Below 5','GEH Below 10','GEH Below 2','GEH Below 5','GEH Below 10'],
     'Modeled Result':[GEHBelow2,GEHBelow5,GEHBelow10,GEHBelow2_pm,GEHBelow5_pm,GEHBelow10_pm]
    }
    FinDat = pd.DataFrame(FinTab)
    FinDat['Modeled Result'] = (FinDat[ 'Modeled Result']*100).apply(lambda x: str(round(x))+'%')
    return(FinDat)
    
#Read Node EVal
#*********************************************************************************
NodeEvalFile_AM = r'./RawVissimOutput/20548_2019_am-existing_V12_Node Results.att'
file = NodeEvalFile_AM
NodeEvalFile_PM = r'./RawVissimOutput/20548_2019_pm-existing_V6_Node Results.att'
VissimDat = ProcessVissimVolumes(file, SimRun = "AVG")
VissimDat_pm = ProcessVissimVolumes(NodeEvalFile_PM, SimRun = "AVG")

#Read Observed TMC Data 
#*********************************************************************************
TMC_dat_am = "CollectedData\AM_tmc_Volumes.csv"
TMC_dat_pm= "CollectedData\PM_tmc_Volumes.csv"
file = TMC_dat_am
ObsDat = ProcessTMC_Data(TMC_dat_am)
ObsDat_pm = ProcessTMC_Data(TMC_dat_pm)

#Merge VISSIM and Observed TMC data
#*********************************************************************************
GEH_Data_am = MergeVISSIM_ObsTMC(VissimDat,ObsDat)
GEH_Data_pm = MergeVISSIM_ObsTMC(VissimDat_pm,ObsDat_pm)

# GEH Table
#*********************************************************************************
FinTab = FinalTable_GEH(GEH_Data_am, GEH_Data_pm)

#Clean up output
#GEH_Data_am = GEH_Data_am.set_index(['NodeName','CardinalDir']).drop(columns='Intersection')
GEH_Data_am = GEH_Data_am.set_index(['Intersection','NodeName','CardinalDir'])
#GEH_Data_pm = GEH_Data_pm.set_index(['NodeName','CardinalDir']).drop(columns='Intersection')
GEH_Data_pm = GEH_Data_pm.set_index(['Intersection','NodeName','CardinalDir'])
#Write to Excel
#*********************************************************************************
OutFi = r'TMC_Calibration.xlsx'
writer=pd.ExcelWriter(OutFi)
GEH_Data_am.to_excel(writer, 'GEH_AM',na_rep='-')
GEH_Data_pm.to_excel(writer, 'GEH_PM',na_rep='-')
FinTab.to_excel(writer, 'Report_Table',na_rep='-')

writer.save() 

    
subprocess.Popen([OutFi],shell=True)  

   

        
    
    
