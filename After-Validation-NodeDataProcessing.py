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

import sys
import os
sys.path.append(os.path.abspath("C:/Users/abibeka/OneDrive - Kittelson & Associates, Inc/Documents/Github/LPGA-VISSIM-DataProcessing"))
from CommonFunctions import ReadExistNodeDelay
from CommonFunctions import ReadMergeVissimObs
from CommonFunctions import ReadDCMINodeDelay

import pandas as pd
import numpy as np
import glob
import subprocess
import re
import seaborn as sns
import matplotlib.pyplot as plt

#Read the Node Results Files
ExistingFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Existing\*Node Results.att')
NoBuildFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\No Build\*Node Results.att')
DCDCFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCDI\*Node Results.att')
DCMIFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCMI\*Node Results.att')
PathToKeyVal = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
KeyValFi = 'NodeEval-KeyValue.xlsx'
KeyValFi = os.path.join(PathToKeyVal,KeyValFi)

match_AM = re.search('_AM',ExistingFile[0])
# Conduct inital cleaning of the VISSIm results and filter it 

PathToKeyVal = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
OutFi = "Report-Vissim-DelayQLenLOS-Results.xlsx"
OutFi = os.path.join(PathToKeyVal,OutFi)
writer=pd.ExcelWriter(OutFi)


i = ExistingFile[1]
VissimDATA = ReadExistNodeDelay(file = i)
file = i
File = KeyValFi
ExistDat= {}
for i in ExistingFile:
    match_AM = re.search('_AM',i)
    if(match_AM):
        ExistDat['AM'] = ReadMergeVissimObs(VissimDATA = ReadExistNodeDelay(file = i),File = KeyValFi,IsFileExisting=True)
        ExistDat['AM'].to_excel(writer,'ExistingAM_DelayRes')
    else:
        ExistDat['PM'] = ReadMergeVissimObs(VissimDATA =ReadExistNodeDelay(file = i),File = KeyValFi,IsFileExisting=True)
        ExistDat['PM'].to_excel(writer,'ExistingPM__DelayRes')
writer.save()

NoBuild = {}
writer=pd.ExcelWriter(OutFi,mode='a')
for i in NoBuildFile:
    match_AM = re.search('_AM',i)
    if(match_AM):
        NoBuild['AM'] = ReadMergeVissimObs(VissimDATA = ReadExistNodeDelay(file = i),File = KeyValFi,SheetNm="NoBuildAM")
        NoBuild['AM'].to_excel(writer,'No-BuildAM_DelayRes')
    else:
        NoBuild['PM'] = ReadMergeVissimObs(VissimDATA =ReadExistNodeDelay(file = i),File = KeyValFi,SheetNm="NoBuildAM")
        NoBuild['PM'].to_excel(writer,'No-BuildPM__DelayRes')
writer.save()    

i = DCDCFile[1]
VissimDATA = ReadExistNodeDelay(file = i)
file = i
File = KeyValFi
#ExistDat= {}

DCDI = {}
writer=pd.ExcelWriter(OutFi,mode='a')
for i in DCDCFile:
    match_AM = re.search('_AM',i)
    if(match_AM):
        DCDI['AM'] = ReadMergeVissimObs(VissimDATA = ReadExistNodeDelay(file = i),File = KeyValFi,SheetNm="DCDI-AM")
        DCDI['AM'].to_excel(writer,'DCDI-AM_DelayRes')
    else:
        DCDI['PM'] = ReadMergeVissimObs(VissimDATA =ReadExistNodeDelay(file = i),File = KeyValFi,SheetNm="DCDI-PM")
        DCDI['PM'].to_excel(writer,'DCDI-PM__DelayRes')
writer.save()    


DCMI = {}
writer=pd.ExcelWriter(OutFi,mode='a')
for i in DCMIFile:
    match_AM = re.search('_AM',i)
    if(match_AM):
        DCMI['AM'] = ReadMergeVissimObs(VissimDATA = ReadDCMINodeDelay(file = i),File = None,SheetNm=None,IsFileDCMI=True)
        DCMI['AM'].to_excel(writer,'DCMI-AM_DelayRes',na_rep=' ')
    else:
        DCMI['PM'] = ReadMergeVissimObs(VissimDATA =ReadDCMINodeDelay(file = i),File = None,SheetNm=None,IsFileDCMI=True)
        DCMI['PM'].to_excel(writer,'DCMI-PM__DelayRes',na_rep=' ')
writer.save()    

        
i = DCMIFile[1]
VissimDATA = ReadExistNodeDelay(ExistingFile[0])
File = KeyValFi




# Create Key Value Pairs for No-Build from Existing
###############################################################################################################
x1 = pd.ExcelFile(KeyValFi)
x1.sheet_names
ExistKeyVal = x1.parse('DCDI-AM')
ExistKeyVal['From']  = ExistKeyVal['From'].str.strip()
ExistKeyVal['To']  = ExistKeyVal['To'].str.strip()
ExistKeyVal['Intersection'] = ExistKeyVal['Intersection'].apply(lambda x: str(x))
DCMCFileKeyVal = ReadDCMINodeDelay(DCMIFile[0])

DCMCFileKeyVal = DCMCFileKeyVal.loc[DCMCFileKeyVal.HourInt=='900-1800',['Intersection','From','To','FromLink','ToLink','Movement']]
lkey_ = ['Intersection','From','To']
DCMCFileKeyVal = DCMCFileKeyVal.merge(ExistKeyVal,left_on =lkey_, right_on=lkey_,how='left')
PathToKeyVal = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
RoughFi = os.path.join(PathToKeyVal,'Rough.csv')
DCMCFileKeyVal.to_csv(RoughFi,na_rep='None')


subprocess.Popen([OutFi],shell=True)  

#Plotting the data
###############################################################################################################
datDict = {'Existing':ExistDat,
           'No-Build': NoBuild,
           'DCDI': DCDI,
           'DCMI':DCMI}
FinDat = pd.DataFrame()
for Scen,dataDict in datDict.items():
    print(Scen)
    for pk,dat in dataDict.items():
        idx = pd.IndexSlice
        dat1 = dat.loc[idx[:,'OverallIntersection'],idx[:,:,'Delay']].droplevel(0,axis=1).droplevel(1).swaplevel(axis=1).stack().reset_index()
        dat1.loc[:,'Scenario'] = Scen
        dat1.loc[:,'Peak'] = pk
        FinDat = pd.concat([dat1,FinDat])

FinDat.HourInt = pd.Categorical(FinDat.HourInt,['900-1800','1800-2700','2700-3600','3600-4500',
                                        '4500-5400','5400-6300','6300-7200','7200-8100',
                                        '8100-9000','9000-9900','9900-10800','10800-11700'])
FinDat = FinDat.sort_values(['HourInt','Intersection'])
FinDat


#https://www.drawingfromdata.com/how-to-rotate-axis-labels-in-seaborn-and-matplotlib
g= sns.catplot(x='HourInt',y='Delay',hue= 'Scenario',col='Intersection',col_wrap=2 ,data= FinDat[FinDat.Peak=='AM'])
for axes in g.axes.flat:
    axes.set_xticklabels(axes.get_xticklabels(), rotation=65, horizontalalignment='right')

g.savefig(PathToKeyVal+'\\TotIntersectionDelay.jpg')