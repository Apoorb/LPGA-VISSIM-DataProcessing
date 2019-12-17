# -*- coding: utf-8 -*-
"""
Created on Fri Sep  6 13:01:32 2019

@author: abibeka
"""

#0.0 Housekeeping. Clear variable space
from IPython import get_ipython  #run magic commands
ipython = get_ipython()
ipython.magic("reset -f")
ipython = get_ipython()


import os
import pandas as pd

#os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\12th-Street-TransitWay\Results')
os.chdir(r'H:\20\20548 - Arlington County Engineering On-Call\009 - 12th Street Transitway Extension\vissim\Results')


x1 = pd.ExcelFile(r'ResultsNameMap.xlsx')
x1.sheet_names
NodeMap = x1.parse('NodeMap')
DirMap = x1.parse('DirMap')


# Read VISSIM File
#******************************************************************************************************************************************************************
def ReadExistNodeDelay(file, NodeMap = NodeMap, DirMap = DirMap):
    '''
    file: Node Evaluation file
    Returns: Cleaned resuls
    Read the Node-Evaluation File
    Ignore the ExistingPMfi. Can be used all AM and PM existing Node Evaluation
    '''
   
    ExistingPMDat=pd.read_csv(file,sep =';',skiprows=1,comment="*")
    ExistingPMDat.columns
    #Use Avg values only
    ExistingPMDat = ExistingPMDat[(ExistingPMDat['$MOVEMENTEVALUATION:SIMRUN'] == 'AVG')]
    #Only use motorized links
    Mask_Not4or5 = ~((ExistingPMDat['MOVEMENT\TOLINK\LINKBEHAVTYPE']==4) | (ExistingPMDat['MOVEMENT\TOLINK\LINKBEHAVTYPE']==5))
    ExistingPMDat = ExistingPMDat[Mask_Not4or5]
    ExistingPMDat.rename(columns={'MOVEMENT':'Mvmt'},inplace=True)
    
    ExPMMvMDat = ExistingPMDat.copy()
    ExPMMvMDat.loc[:,"HourInt"] = 'Nan'
    ExPMMvMDat.loc[:,'HourInt'] = ExPMMvMDat.TIMEINT
    ExPMMvMDat.HourInt = pd.Categorical(ExPMMvMDat.HourInt,['900-4500'])
    ExPMMvMDat.rename(columns= {'VEHS(ALL)':'veh','QLEN':'Qlen','QLENMAX':'QLenMax','VEHDELAY(ALL)':'Delay','MOVEMENT\DIRECTION':'VissimDir'},inplace=True)
    new = ExPMMvMDat.Mvmt.str.split(':',expand=True)
    ExPMMvMDat['Intersection'] = (new[0].str.split('-',n=1,expand=True)[0]).astype('int')
    ExPMMvMDat = ExPMMvMDat[['Intersection','VissimDir','veh','Qlen','QLenMax','Delay']]
    
    #Handle duplicate rows 
    ExPMMvMDat.loc[:,'Delay'] = ExPMMvMDat.Delay * ExPMMvMDat.veh
    ExPMMvMDat = ExPMMvMDat.groupby(['Intersection','VissimDir']).agg({'veh':'sum', 'Delay':'sum','Qlen':'min','QLenMax':'min'}).reset_index()
    ExPMMvMDat.loc[:,'Delay'] = ExPMMvMDat.Delay / ExPMMvMDat.veh
    
    ExPMMvMDat.loc[:,'LOS'] = ExPMMvMDat.Delay.apply(LOS_Calc)
    ExPMMvMDat = ExPMMvMDat.merge(DirMap, left_on = 'VissimDir',right_on = 'VissimDir',how = 'left') 
    #--------------------------------------------------------------------------------------------------
    ExPMMvMDat = pd.merge(ExPMMvMDat,NodeMap,left_on=['Intersection'],
                 right_on=["NodeNum"], how = 'left')
    ExPMMvMDat.CardinalDir = pd.Categorical(ExPMMvMDat.CardinalDir,[
        'EBL','EBT','EBR','WBL','WBT','WBR','NBL','NBT','NBR','SBL','SBT','SBR','Intersection','HSep'
])
    ExPMMvMDat = ExPMMvMDat[['NodeNum','CardinalDir','veh','Delay','LOS','Qlen','QLenMax']]
    ExPMMvMDat.columns
    ExPMMvMDat =ExPMMvMDat[~ExPMMvMDat.CardinalDir.isna()]
    # Remove a duplicate right turn for Node 63 and EBR . Will deal with it later
#    ExPMMvMDat = ExPMMvMDat[~ExPMMvMDat.duplicated(subset=['NodeNum','CardinalDir'],keep='first')]
    #Handled above
    ExPMMvMDat = ExPMMvMDat.pivot(index ='NodeNum',columns = 'CardinalDir', values= ['veh','Delay','LOS','Qlen','QLenMax'])
    ExPMMvMDat.columns = ExPMMvMDat.columns.swaplevel(0, 1)
    ExPMMvMDat = ExPMMvMDat.stack()
    ExPMMvMDat= ExPMMvMDat[['EBL','EBT','EBR','WBL','WBT','WBR','NBL','NBT','NBR','SBL','SBT','SBR','Intersection']]
    
    IntLevel = ExPMMvMDat.index.levels[0]
    PM_level = ExPMMvMDat.index.levels[1]
    mux = pd.MultiIndex.from_product([IntLevel,['veh','Delay','LOS','Qlen','QLenMax']], names=['Intersection','PMs'])
    ExPMMvMDat = ExPMMvMDat.reindex(mux,axis=0)
    return(ExPMMvMDat)
    
#******************************************************************************************************************************************************************
def LOS_Calc(Delay):
    LOS = 'Z'
    if Delay<=10:
        LOS = 'A'
    elif ((Delay>10) & (Delay<=20)):
        LOS = 'B'
    elif ((Delay>20) & (Delay<=35)):
        LOS = 'C'
    elif ((Delay>35) & (Delay<=55)):
        LOS = 'D'
    elif ((Delay>55) & (Delay<=80)):
        LOS = 'E'
    elif Delay > 80:
        LOS = 'F'
    else:
        LOS = 'Z'
    return(LOS)
    
    
#******************************************************************************************************************************************************************
NodeEvalFile_AM = r'./RawVissimOutput/20548_2019_am-existing_V5_Node Results.att'
OutFi = r'PythonNodeEvalRes.xlsx'
file = NodeEvalFile_AM

writer=pd.ExcelWriter(OutFi)
Dat = ReadExistNodeDelay(file)
Dat.to_excel(writer, 'Existing',na_rep='-')
writer.save() 

    
    
   

        
    
    
