# -*- coding: utf-8 -*-
"""
Created on Tue Sep 10 13:25:02 2019

@author: abibeka
"""


#0.0 Housekeeping. Clear variable space
from IPython import get_ipython  #run magic commands
ipython = get_ipython()
ipython.magic("reset -f")
ipython = get_ipython()

import sys
import os
sys.path.append(os.path.abspath("C:/Users/abibeka/OneDrive - Kittelson & Associates, Inc/Documents/Github/LPGA-VISSIM-DataProcessing"))

import pandas as pd
import numpy as np
import glob
import subprocess
import re


#Read the Node Results Files
ExistAM=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Existing\*AM*_Vehicle Network Performance Evaluation Results.att')
ExistPM=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Existing\*PM*_Vehicle Network Performance Evaluation Results.att')


NoBuildAM=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\No Build\*AM*_Vehicle Network Performance Evaluation Results.att')
NoBuildPM=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\No Build\*PM*_Vehicle Network Performance Evaluation Results.att')

DCDI_AM=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCDI\*AM*_Vehicle Network Performance Evaluation Results.att')
DCDI_PM=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCDI\*PM*_Vehicle Network Performance Evaluation Results.att')

DCMI_AM=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCMI\*AM*_Vehicle Network Performance Evaluation Results.att')
DCMI_PM=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCMI\*PM*_Vehicle Network Performance Evaluation Results.att')

file = ExistAM[0]

def CleanNetPerFun(file):
    '''
  
    '''
    # Get the time keys to convert Vissim intervals to Hours of the day
    PathToVis = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
    TimeConvFi  = 'TravelTimeKeyValuePairs.xlsx'
    TimeConvFi = pd.ExcelFile(os.path.join(PathToVis,TimeConvFi))
    TimeConvFi.sheet_names
    TimeKeys = TimeConvFi.parse('TimeVissimKey')
    TimeKeys_AM = TimeKeys.loc[TimeKeys.AnalysisPeriod =='AM',['Time','TIMEINT']]
    TimeKeys_PM = TimeKeys.loc[TimeKeys.AnalysisPeriod =='PM',['Time','TIMEINT']]

    Dat = pd.read_csv(file,sep =';',skiprows=62)
    Dat.columns
    Dat = Dat[Dat['$VEHICLENETWORKPERFORMANCEMEASUREMENTEVALUATION:SIMRUN']=='AVG']
    Dat.columns =[x.strip() for x in Dat.columns]
    Dat = Dat.rename(columns = {
                          'DELAYAVG(ALL)': 'Average Delay — All (sec)',
                          'DELAYTOT(ALL)': 'Total Delay — All (sec)',
                          'STOPSTOT(ALL)': 'Total Stops — All',
                          'STOPSAVG(ALL)': 'Average Stops — All',
                          'SPEEDAVG(ALL)': 'Average Speed — All',
                          'DELAYLATENT' : 'Latent Delay (sec)',
                          'DEMANDLATENT': 'Latent Demand (veh)'
                          })
    Dat = Dat[['TIMEINT','Average Delay — All (sec)', 
    'Total Delay — All (sec)','Total Stops — All','Average Stops — All',
    'Average Speed — All','Latent Delay (sec)', 'Latent Demand (veh)']]   
    Dat.dtypes
    listVar = ['Average Delay — All (sec)', 
    'Total Delay — All (sec)','Average Stops — All',
    'Average Speed — All','Latent Delay (sec)', 'Latent Demand (veh)']
    Dat.loc[:,listVar] = Dat.loc[:,listVar].applymap(lambda x: format(x,',.1f'))
    Dat.loc[:,'Total Stops — All'] =Dat.loc[:,'Total Stops — All'].apply(lambda x: format(x,','))
    match_AM = re.search('_AM',file)
    if match_AM:
        Dat = TimeKeys_AM.merge(Dat,left_on='TIMEINT',right_on='TIMEINT',how='right').drop(columns='TIMEINT')
    else:
        Dat = TimeKeys_PM.merge(Dat,left_on='TIMEINT',right_on='TIMEINT',how='right').drop(columns='TIMEINT')
    return(Dat)

DataDict = {
        'ExistAM': ExistAM[0],
        'ExistPM': ExistPM[0],
        'NoBuildAM': NoBuildAM[0],
        'NoBuildPM': NoBuildPM[0],
        'DCDI_AM': DCDI_AM[0],
        'DCDI_PM': DCDI_PM[0],
        'DCMI_AM': DCMI_AM[0],
        'DCMI_PM': DCMI_PM[0],
        }


OutFi = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\NetworkPerfEvalRes.xlsx'

writer=pd.ExcelWriter(OutFi)
for name, file, in DataDict.items():
    Dat = CleanNetPerFun(file)
    Dat.to_excel(writer,name,na_rep=' ')
writer.save()    

        