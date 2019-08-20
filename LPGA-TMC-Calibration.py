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
import datetime

PathToExist = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Existing'
ExistingPMfi = '20834_Existing_PM--C1C2C3C4C5C6C7_Node Results.att'
ExistingPMfi = os.path.join(PathToExist,ExistingPMfi)
ExistingAMfi ='20834_Existing_PM--C1C2C3C4C5C6C7_Node Results.att'
ExistingAMfi = os.path.join(PathToExist,ExistingAMfi)


# VISSIM File
#*********************************************************************************
ExistingPMDat=pd.read_csv(ExistingPMfi,sep =';',skiprows=28)
ExistingPMDat.columns

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

ExPMMvMDat.dtypes

PathToKeyVal = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
KeyValFi = 'NodeEval-KeyValue.xlsx'
KeyValFi = os.path.join(PathToKeyVal,KeyValFi)

x1 = pd.ExcelFile(KeyValFi)
x1.sheet_names
ExistPM_Vissim = x1.parse('ExistingAM')
ExistPM_Vissim['From']  = ExistPM_Vissim['From'].str.strip()
ExistPM_Vissim['To']  = ExistPM_Vissim['To'].str.strip()
ExistPM_Vissim['Intersection'] = ExistPM_Vissim['Intersection'].apply(lambda x: str(x))
ExistPM_Vissim = pd.merge(ExistPM_Vissim,ExPMMvMDat,left_on=['Intersection','From','To'],
         right_on=['Intersection','From','To'], how = 'left')
ExistPM_Vissim.Movement = ExistPM_Vissim.Movement.str.strip()
ExistPM_Vissim.Movement = pd.Categorical(ExistPM_Vissim.Movement,[
        'EB U/L','EB L/T','EB T', 'EB R'
,'WB L','WB U/L', 'WB T/R','WB T', 'WB R','NB U/L', 'NB L','NB T/R',
 'NB T', 'NB R','SB U/L','SB L', 'SB T','SB R','SB T/R','SB L/T/R'
])

ExistPM_Vissim.HourInt = pd.Categorical(ExistPM_Vissim.HourInt,['900-4500','4500-8100','8100-11700'])

ExistPM_Vissim = ExistPM_Vissim.sort_values(['Intersection','HourInt','Movement'])

ExistPM_Vissim.groupby(['Intersection','HourInt','Movement'])['VEHS(ALL)'].sum()