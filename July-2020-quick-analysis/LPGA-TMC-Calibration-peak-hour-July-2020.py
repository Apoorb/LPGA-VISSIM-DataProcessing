# -*- coding: utf-8 -*-
"""
Created on Mon Aug 19 13:47:01 2019

@author: abibeka
Read data from Node Data files
Get the GEH for every 15 mins - in the 3 hour analysis period.
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
import subprocess


PathToExist = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\July-6-2020\Models---July-2020\VISSIM\Existing'
ExistingPMfi = '20834_Existing_PM_Node Results.att'
ExistingPMfi = os.path.join(PathToExist,ExistingPMfi)

# Read VISSIM File
#*********************************************************************************
def ReadExistNodeRes(file =ExistingPMfi):
    ExistingPMDat=pd.read_csv(file,sep =';',skiprows=28)
    ExistingPMDat.columns
    #Use Avg values only
    ExistingPMDat = ExistingPMDat[(ExistingPMDat['$MOVEMENTEVALUATION:SIMRUN'] == 'AVG')]
    ExistingPMDat.rename(columns={'MOVEMENT':'Mvmt'},inplace=True)
    
    # Use the data for movements only. Ignore Intersection summary for now
    ExPMMvMDat = ExistingPMDat[ExistingPMDat.Mvmt.str.len()>1]
    
    ExPMMvMDat.loc[:,"HourInt"] = 'Nan'
            
#    FullHr = ExPMMvMDat.TIMEINT.isin(['900-1800','1800-2700','2700-3600','3600-4500',
#                                        '4500-5400','5400-6300','6300-7200','7200-8100',
#                                        '8100-9000','9000-9900','9900-10800','10800-11700'])
    ExPMMvMDat.loc[:,'HourInt'] = ExPMMvMDat.TIMEINT
    ExPMMvMDat = ExPMMvMDat.groupby(['Mvmt','HourInt'])['VEHS(ALL)'].sum().reset_index()
    new = ExPMMvMDat.Mvmt.str.split(':',expand=True)
    #Correct some minor stuff from Existing PM
    new.loc[new[1] == ' WB LPGA Blvd@9164.7',2] = ' WB LPGA Blvd@9164.7'
    new.loc[new[1] == ' WB LPGA Blvd@9164.7',1] = 'None'
    ExPMMvMDat['Intersection'] = new[0].str.split('-',n=1,expand=True)[0]
    ExPMMvMDat['From'] = new[1].str.split('@',n=1,expand=True)[0]
    ExPMMvMDat.loc[:,'To'] = new[2].str.split('@',n=1,expand=True)[0]
    ExPMMvMDat['From']  = ExPMMvMDat['From'].str.strip()
    ExPMMvMDat['To']  = ExPMMvMDat['To'].str.strip()
    ExPMMvMDat.Intersection = ExPMMvMDat.Intersection.str.strip()
    return(ExPMMvMDat)



ExPMMvMDat = ReadExistNodeRes(ExistingPMfi)
ExistingAMfi ='20834_Existing_AM_Node Results.att'
ExistingAMfi = os.path.join(PathToExist,ExistingAMfi)
ExAMMvMDat = ReadExistNodeRes(ExistingAMfi)

# Get the Key Value pair - for report
PathToKeyVal = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
KeyValFi = 'NodeEval-KeyValue.xlsx'
KeyValFi = os.path.join(PathToKeyVal,KeyValFi)


PMConvFactors = {
 '900-1800':.88,
 '1800-2700':.96,
 '2700-3600':.96,
 '3600-4500':.97,
 '4500-5400':.98,
 '5400-6300':.95,
 '6300-7200':1.05,
 '7200-8100':1.02,
 '8100-9000':.94,
 '9000-9900':.79,
 '9900-10800':.71,
 '10800-11700':.71
 }


AMConvFactors = {
 '900-1800':.38,
 '1800-2700':.54,
 '2700-3600':.60,
 '3600-4500':.66,
 '4500-5400':.90,
 '5400-6300':1.12,
 '6300-7200':1.12,
 '7200-8100':0.86,
 '8100-9000':.82,
 '9000-9900':.82,
 '9900-10800':.74,
 '10800-11700':.66
 }

def ReadMergeVissimObs(VissimDATA = ExPMMvMDat, File = KeyValFi,VolTime='PeakPMVol' , Peak = (.98+.95+1.05+1.02), 
    ConversionFactors = PMConvFactors):
    '''
    Default is setup for Existing PM Analysis
    VissimDATA = ExPMMvMDat,  # Vissim results to merge
    File = KeyValFi, # Key value file for PM
    VolTime='PeakPMVol' ,  # Use PM obs vol by default
    Peak = (.98+.95+1.05+1.02),  # PM values
    ConversionFactors = Converting vol to 15 min
    '''
    x1 = pd.ExcelFile(File)
    x1.sheet_names
    ExistPM_Vissim = x1.parse('ExistingAM')
    ExistPM_Vissim['From']  = ExistPM_Vissim['From'].str.strip()
    ExistPM_Vissim['To']  = ExistPM_Vissim['To'].str.strip()
    ExistPM_Vissim['Intersection'] = ExistPM_Vissim['Intersection'].apply(lambda x: str(x))
    ExistPM_Vissim = pd.merge(ExistPM_Vissim,VissimDATA,left_on=['Intersection','From','To'],
             right_on=['Intersection','From','To'], how = 'left')
    ExistPM_Vissim.Movement = ExistPM_Vissim.Movement.str.strip()
    ExistPM_Vissim.Movement = pd.Categorical(ExistPM_Vissim.Movement,[
            'EB U/L','EB L/T','EB T', 'EB R'
    ,'WB L','WB U/L', 'WB T/R','WB T', 'WB R','NB U/L', 'NB L','NB T/R',
     'NB T', 'NB R','SB U/L','SB L', 'SB T','SB R','SB T/R','SB L/T/R'
    ])
    ExistPM_Vissim.HourInt = pd.Categorical(ExistPM_Vissim.HourInt,['900-1800','1800-2700','2700-3600','3600-4500',
                                        '4500-5400','5400-6300','6300-7200','7200-8100',
                                        '8100-9000','9000-9900','9900-10800','10800-11700'])
    ExistPM_Vissim = ExistPM_Vissim.sort_values(['Intersection','HourInt','Movement'])
    ExistPM_Vissim = ExistPM_Vissim.groupby(['Intersection','HourInt','Movement'])['VEHS(ALL)'].sum().reset_index()
    ExistPM_Vissim.rename(columns= {'VEHS(ALL)':'VissimVol'},inplace=True)
    # Get the observed TMC data
    ObsTMCDat = x1.parse('ObsTMC',usecols=['Intersection','Movement',VolTime])
    ObsTMCDat.Intersection = ObsTMCDat.Intersection.apply(lambda x: str(x))
    ObsTMCDat.loc[:,'HourInt'] = '4500-8100'
    ObsTMCDat.Movement = ObsTMCDat.Movement.str.strip()
    ObsTMCDat.rename(columns = {VolTime:'ObsVol'},inplace = True)
    ObsTMCDat.ObsVol = ObsTMCDat.ObsVol/Peak
    L =[]
    for key,val in ConversionFactors.items():
        Temp = ObsTMCDat.copy()
        Temp.loc[:,'HourInt'] = key
        Temp.ObsVol = Temp.ObsVol * val
        L.append(Temp)
    ObsTMCDat = pd.concat(L).reset_index(drop=True)    
    ObsTMCDat = ObsTMCDat.groupby(['Intersection','Movement','HourInt'])['ObsVol'].sum().reset_index()

    # Merge with VISSIM
    ExistPM_Vissim = pd.merge(ExistPM_Vissim,ObsTMCDat,
             left_on =['Intersection','HourInt','Movement'],
             right_on = ['Intersection','HourInt','Movement'],
             how='left')
    ExistPM_Vissim=\
        ExistPM_Vissim.query("HourInt.isin(['4500-5400','5400-6300','6300-7200','7200-8100'])")
    Debug = ExistPM_Vissim
    Debug.loc[:,'Num'] = ((ExistPM_Vissim.VissimVol - ExistPM_Vissim.ObsVol)**2)
    Debug.loc[:,'Denom'] = ((ExistPM_Vissim.VissimVol + ExistPM_Vissim.ObsVol)/2)
    Debug.loc[:,'Ratio'] = Debug.loc[:,'Num']/  Debug.loc[:,'Denom']
    Debug.loc[:,'Ratio_Sqrt'] = np.sqrt(Debug.loc[:,'Ratio'])
    ExistPM_Vissim.loc[:,'GEH'] = np.sqrt(((ExistPM_Vissim.VissimVol - ExistPM_Vissim.ObsVol)**2)/ ((ExistPM_Vissim.VissimVol + ExistPM_Vissim.ObsVol)/2))
    return(ExistPM_Vissim)

PathToKeyVal = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
KeyValFi = 'NodeEval-KeyValue.xlsx'
KeyValFi = os.path.join(PathToKeyVal,KeyValFi)

ExistPM_Vissim = ReadMergeVissimObs()
ExistAM_Vissim = ReadMergeVissimObs(VissimDATA = ExAMMvMDat, File = KeyValFi,VolTime='PeakAMVol' ,
                       Peak = (.90 + 1.12 + 1.12 + 0.86), 
                       ConversionFactors = AMConvFactors)
ExistAM_Vissim.set_index('Intersection',inplace=True)
IntersectionKey = {'1':'LPGA/Tomoka Farms',
                    '2':'LPGA/I-95 SB Ramp',
                    '3':'LPGA/I-95 NB Ramp',
                    '4':'LPGA/Technology',
                    '5':'LPGA/Williamson',
                    '6':'LPGA/Clyde Morris'}
ExistAM_Vissim =ExistAM_Vissim.rename(axis = 0, mapper=IntersectionKey).reset_index()
ExistAM_Vissim = ExistAM_Vissim[['Intersection','Movement','HourInt','ObsVol','VissimVol','GEH']]
ExistPM_Vissim = ExistPM_Vissim[['Intersection','Movement','HourInt','ObsVol','VissimVol','GEH']]

##########################
##########################
# HACK
#########################
path_to_i95_tmc_dir = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\July-6-2020\Models---July-2020\VISSIM'
file_nm  = "i-95-processed-vissim-observed-volumes.xlsx"
path_to_i95_tmc_file = os.path.join(path_to_i95_tmc_dir,file_nm)
x1 = pd.ExcelFile(path_to_i95_tmc_file)
x1.sheet_names
ExistAM_Vissim_i95 = x1.parse('am_i95_volume',index_col=0)
ExistPM_Vissim_i95 = x1.parse('pm_i95_volume',index_col=0)

ExistAM_Vissim = pd.concat([ExistAM_Vissim,ExistAM_Vissim_i95])
ExistPM_Vissim = pd.concat([ExistPM_Vissim,ExistPM_Vissim_i95])
##########################
# HACK END
#########################
ExistAM_Vissim.Intersection = pd.Categorical(ExistAM_Vissim.Intersection,[
    'LPGA/Tomoka Farms',
    'LPGA/I-95 SB Ramp',
    'LPGA/I-95 NB Ramp',
    'LPGA/Technology',
    'LPGA/Williamson',
    'LPGA/Clyde Morris',
    'SB I-95 (SR40 to SB OffRamp)',
    'SB I-95 (SB OffRamp to SB LoopRamp)',
    'SB I-95 (SB LoopRamp to SB On-Ramp)',
    'SB I-95 (SB On-Ramp to US92)',
    'NB I-95 (US92 to NB OffRamp)',
    'NB I-95 (NB OffRamp to NB LoopRamp)',
    'NB I-95 ( NB LoopRamp to NB On-Ramp)',
    'NB I-95 (NB On-Ramp to SR40)'
])


ExistAM_Vissim = ExistAM_Vissim.set_index(['Intersection','Movement','HourInt']).unstack()
ExistAM_Vissim.loc[:,:'ObsVol'] = ExistAM_Vissim.loc[:,:'ObsVol'].round()
ExistAM_Vissim.loc[:,:'VissimVol'] = ExistAM_Vissim.loc[:,:'VissimVol'].round()
ExistAM_Vissim.loc[:,:'GEH'] = ExistAM_Vissim.loc[:,:'GEH'].round(2)
ExistAM_Vissim.columns = ExistAM_Vissim.columns.swaplevel(0, 1)
mux = pd.MultiIndex.from_product([['4500-5400','5400-6300','6300-7200','7200-8100'],
                                  [u'ObsVol',u'VissimVol',u'GEH']], names=ExistAM_Vissim.index.names)
ExistAM_Vissim = ExistAM_Vissim.reindex(mux,axis=1)


ExistPM_Vissim.set_index('Intersection',inplace=True)
ExistPM_Vissim =ExistPM_Vissim.rename(axis = 0, mapper=IntersectionKey).reset_index()
ExistPM_Vissim.Intersection = pd.Categorical(ExistPM_Vissim.Intersection,[
    'LPGA/Tomoka Farms',
    'LPGA/I-95 SB Ramp',
    'LPGA/I-95 NB Ramp',
    'LPGA/Technology',
    'LPGA/Williamson',
    'LPGA/Clyde Morris',
    'SB I-95 (SR40 to SB OffRamp)',
    'SB I-95 (SB OffRamp to SB LoopRamp)',
    'SB I-95 (SB LoopRamp to SB On-Ramp)',
    'SB I-95 (SB On-Ramp to US92)',
    'NB I-95 (US92 to NB OffRamp)',
    'NB I-95 (NB OffRamp to NB LoopRamp)',
    'NB I-95 ( NB LoopRamp to NB On-Ramp)',
    'NB I-95 (NB On-Ramp to SR40)'
])

ExistPM_Vissim = ExistPM_Vissim.set_index(['Intersection','Movement','HourInt']).unstack()
ExistPM_Vissim.loc[:,:'ObsVol'] = ExistPM_Vissim.loc[:,:'ObsVol'].round()
ExistPM_Vissim.loc[:,:'VissimVol'] = ExistPM_Vissim.loc[:,:'VissimVol'].round()
ExistPM_Vissim.loc[:,:'GEH'] = ExistPM_Vissim.loc[:,:'GEH'].round(2)
ExistPM_Vissim.columns = ExistPM_Vissim.columns.swaplevel(0, 1)
ExistPM_Vissim = ExistPM_Vissim.reindex(mux,axis=1)
ExistAM_Vissim = ExistAM_Vissim.dropna(axis=0)
ExistPM_Vissim = ExistPM_Vissim.dropna(axis=0)

#Get Final Table
idx = pd.IndexSlice
TotalValuesPM = ExistPM_Vissim.loc[:,idx[:,'GEH']].shape[0]* ExistPM_Vissim.loc[:,idx[:,'GEH']].shape[1]
GEHBelow2= sum((ExistPM_Vissim.loc[:,idx[:,'GEH']]<=2).values.ravel())/TotalValuesPM
GEHBelow5= sum((ExistPM_Vissim.loc[:,idx[:,'GEH']]<=5).values.ravel())/TotalValuesPM
GEHBelow10= sum((ExistPM_Vissim.loc[:,idx[:,'GEH']]<=10).values.ravel())/TotalValuesPM

TotalValuesAM = ExistAM_Vissim.loc[:,idx[:,'GEH']].shape[0]*ExistAM_Vissim.loc[:,idx[:,'GEH']].shape[1]
GEHBelow2_am =sum((ExistAM_Vissim.loc[:,idx[:,'GEH']]<=2).values.ravel())/TotalValuesAM
GEHBelow5_am = sum((ExistAM_Vissim.loc[:,idx[:,'GEH']]<=5).values.ravel())/TotalValuesAM
GEHBelow10_am = sum((ExistAM_Vissim.loc[:,idx[:,'GEH']]<=10).values.ravel())/TotalValuesAM

FinTab = {'Time' : ['AM Peak Hour','AM Peak Hour','AM Peak Hour','PM Peak Hour','PM Peak Hour','PM Peak Hour'],
'Calibration Metric' : ['GEH Below 2','GEH Below 5','GEH Below 10','GEH Below 2','GEH Below 5','GEH Below 10'],
 'Modeled Result':[GEHBelow2_am,GEHBelow5_am,GEHBelow10,GEHBelow2,GEHBelow5,GEHBelow10]
}

FinDat = pd.DataFrame(FinTab)
FinDat['Modeled Result'] = (FinDat[ 'Modeled Result']*100).apply(lambda x: str(round(x))+'%')

PathToKeyVal = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\July-6-2020\Models---July-2020\VISSIM'
OutFi = "Report-TT-GEH-Results.xlsx"
OutFi = os.path.join(PathToKeyVal,OutFi)
writer=pd.ExcelWriter(OutFi,mode ='a')
FinDat.to_excel(writer, 'FinalRes')
ExistAM_Vissim.to_excel(writer,'ExistingAM_GEH')
ExistPM_Vissim.to_excel(writer,'ExistingPM_GEH')
writer.save()

subprocess.Popen([OutFi],shell=True)  