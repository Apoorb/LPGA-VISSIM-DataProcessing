# -*- coding: utf-8 -*-
"""
Created on Mon Aug 19 13:47:01 2019

@author: abibeka
Read data from Node Data files
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


PathToExist = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Existing'
ExistingPMfi = '20834_Existing_PM--C1C2C3C4C5C6C7_Node Results.att'
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
            
    BeforePeak = ExPMMvMDat.TIMEINT.isin(['900-1800','1800-2700','2700-3600','3600-4500'])
    PeakHour = ExPMMvMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100'])
    AfterPeak = ExPMMvMDat.TIMEINT.isin(['8100-9000','9000-9900','9900-10800','10800-11700'])
    BeforePeak.sum()
    PeakHour.sum()
    AfterPeak.sum()
    ExPMMvMDat.loc[BeforePeak,'HourInt'] = '900-4500'
    ExPMMvMDat.loc[PeakHour,'HourInt'] = '4500-8100'
    ExPMMvMDat.loc[AfterPeak,'HourInt'] = '8100-11700'
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
ExistingAMfi ='20834_Existing_AM--C1C2aC3C4C5C6C7C8_Node Results.att'
ExistingAMfi = os.path.join(PathToExist,ExistingAMfi)
ExAMMvMDat = ReadExistNodeRes(ExistingAMfi)

# Get the Key Value pair - for report
PathToKeyVal = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
KeyValFi = 'NodeEval-KeyValue.xlsx'
KeyValFi = os.path.join(PathToKeyVal,KeyValFi)

def ReadMergeVissimObs(VissimDATA = ExPMMvMDat, File = KeyValFi,VolTime='PeakPMVol' , Peak = (.98+.95+1.05+1.02), BeforePeak = (.88+.96+.96+.97),
    AfterPeak = (.94+.79+.71+.71)):
    '''
    Default is setup for Existing PM Analysis
    VissimDATA = ExPMMvMDat,  # Vissim results to merge
    File = KeyValFi, # Key value file for PM
    VolTime='PeakPMVol' ,  # Use PM obs vol by default
    Peak = (.98+.95+1.05+1.02),  # PM values
    BeforePeak = (.88+.96+.96+.97), # PM values
    AfterPeak = (.94+.79+.71+.71)) # PM values
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
    ExistPM_Vissim.HourInt = pd.Categorical(ExistPM_Vissim.HourInt,['900-4500','4500-8100','8100-11700'])
    ExistPM_Vissim = ExistPM_Vissim.sort_values(['Intersection','HourInt','Movement'])
    ExistPM_Vissim = ExistPM_Vissim.groupby(['Intersection','HourInt','Movement'])['VEHS(ALL)'].sum().reset_index()
    ExistPM_Vissim.rename(columns= {'VEHS(ALL)':'VissimVol'},inplace=True)
    # Get the observed TMC data
    ObsTMCDat = x1.parse('ObsTMC',usecols=['Intersection','Movement',VolTime])
    ObsTMCDat.Intersection = ObsTMCDat.Intersection.apply(lambda x: str(x))
    ObsTMCDat.loc[:,'HourInt'] = '4500-8100'
    ObsTMCDat.Movement = ObsTMCDat.Movement.str.strip()
    ObsTMCDat.rename(columns = {VolTime:'ObsVol'},inplace = True)
    #Fill Shoulder Hour Volumes
    Temp,Temp2 = ObsTMCDat.copy(),ObsTMCDat.copy()
    Temp.loc[:,'HourInt'] = '900-4500'
    Temp.ObsVol = Temp.ObsVol * BeforePeak/Peak
    Temp2.loc[:,'HourInt'] = '8100-11700'
    Temp2.ObsVol = Temp2.ObsVol * AfterPeak/Peak
    # GEt the full data
    ObsTMCDat = pd.concat([Temp,ObsTMCDat,Temp2]).reset_index(drop=True)
    
    # Merge with VISSIM
    ExistPM_Vissim = pd.merge(ExistPM_Vissim,ObsTMCDat,
             left_on =['Intersection','HourInt','Movement'],
             right_on = ['Intersection','HourInt','Movement'],
             how='left')
    ExistPM_Vissim.ObsVol = ExistPM_Vissim.ObsVol.apply(lambda x: np.ceil(x))
    ExistPM_Vissim.loc[:,'GEH'] = np.sqrt(((ExistPM_Vissim.VissimVol - ExistPM_Vissim.ObsVol)**2)/ ((ExistPM_Vissim.VissimVol + ExistPM_Vissim.ObsVol)/2))
    return(ExistPM_Vissim)

PathToKeyVal = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
KeyValFi = 'NodeEval-KeyValue.xlsx'
KeyValFi = os.path.join(PathToKeyVal,KeyValFi)

ExistPM_Vissim = ReadMergeVissimObs()
ExistAM_Vissim = ReadMergeVissimObs(VissimDATA = ExAMMvMDat, File = KeyValFi,VolTime='PeakAMVol' ,
                       BeforePeak = (.60+.71+.78+.81),
                       Peak = (.94 + 1.05 + 1.04 + 0.96), 
                       AfterPeak = (.92+.88+.77+.69))

ExistAM_Vissim.set_index('Intersection',inplace=True)
IntersectionKey = {'1':'LPGA/Tomoka Farms',
                    '2':'LPGA/I-95 SB Ramp',
                    '3':'LPGA/I-95 NB Ramp',
                    '4':'LPGA/Technology',
                    '5':'LPGA/Williamson',
                    '6':'LPGA/Clyde Morris'}
ExistAM_Vissim =ExistAM_Vissim.rename(axis = 0, mapper=IntersectionKey).reset_index()
ExistAM_Vissim = ExistAM_Vissim.set_index(['Intersection','Movement','HourInt']).unstack()
ExistAM_Vissim.loc[:,:'ObsVol'] = ExistAM_Vissim.loc[:,:'ObsVol'].round()
ExistAM_Vissim.loc[:,:'VissimVol'] = ExistAM_Vissim.loc[:,:'VissimVol'].round()
ExistAM_Vissim.loc[:,:'GEH'] = ExistAM_Vissim.loc[:,:'GEH'].round(2)
ExistAM_Vissim.columns = ExistAM_Vissim.columns.swaplevel(0, 1)
mux = pd.MultiIndex.from_product([[u'900-4500',u'4500-8100', u'8100-11700'],
                                  [u'ObsVol',u'VissimVol',u'GEH']], names=ExistAM_Vissim.index.names)
ExistAM_Vissim = ExistAM_Vissim.reindex(mux,axis=1)


ExistPM_Vissim.set_index('Intersection',inplace=True)
ExistPM_Vissim =ExistPM_Vissim.rename(axis = 0, mapper=IntersectionKey).reset_index()
ExistPM_Vissim = ExistPM_Vissim.set_index(['Intersection','Movement','HourInt']).unstack()
ExistPM_Vissim.loc[:,:'ObsVol'] = ExistPM_Vissim.loc[:,:'ObsVol'].round()
ExistPM_Vissim.loc[:,:'VissimVol'] = ExistPM_Vissim.loc[:,:'VissimVol'].round()
ExistPM_Vissim.loc[:,:'GEH'] = ExistPM_Vissim.loc[:,:'GEH'].round(2)
ExistPM_Vissim.columns = ExistPM_Vissim.columns.swaplevel(0, 1)

ExistPM_Vissim = ExistPM_Vissim.reindex(mux,axis=1)

PathToKeyVal = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
OutFi = "Report-GEH-Results.xlsx"
OutFi = os.path.join(PathToKeyVal,OutFi)


writer=pd.ExcelWriter(OutFi,mode = 'a')
ExistAM_Vissim.to_excel(writer,'ExistingAM_GEH')
ExistPM_Vissim.to_excel(writer,'ExistingPM_GEH')
writer.save()

subprocess.Popen([OutFi],shell=True)  
idx = pd.IndexSlice
TotalValuesPM = ExistPM_Vissim.loc[:,idx[:,'GEH']].shape[0]* ExistPM_Vissim.loc[:,idx[:,'GEH']].shape[1]
GEHBelow2= sum((ExistPM_Vissim.loc[:,idx[:,'GEH']]<=2).values.ravel())/TotalValuesPM
GEHBelow5= sum((ExistPM_Vissim.loc[:,idx[:,'GEH']]<=5).values.ravel())/TotalValuesPM
GEHBelow10= sum((ExistPM_Vissim.loc[:,idx[:,'GEH']]<=10).values.ravel())/TotalValuesPM

TotalValuesAM = ExistAM_Vissim.loc[:,idx[:,'GEH']].shape[0]*ExistAM_Vissim.loc[:,idx[:,'GEH']].shape[1]
GEHBelow2_am =sum((ExistAM_Vissim.loc[:,idx[:,'GEH']]<=2).values.ravel())/TotalValuesAM
GEHBelow5_am = sum((ExistAM_Vissim.loc[:,idx[:,'GEH']]<=5).values.ravel())/TotalValuesAM
GEHBelow10_am = sum((ExistAM_Vissim.loc[:,idx[:,'GEH']]<=10).values.ravel())/TotalValuesAM


