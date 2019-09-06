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

import pandas as pd
import numpy as np
import glob
import subprocess
import re

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

ExistDat= {}
for i in ExistingFile:
    match_AM = re.search('_AM',i)
    if(match_AM):
        ExistDat['AM'] = ReadMergeVissimObs(VissimDATA = ReadExistNodeDelay(file = i),File = KeyValFi)
        ExistDat['AM'].to_excel(writer,'ExistingAM_DelayRes')
    else:
        ExistDat['PM'] = ReadMergeVissimObs(VissimDATA =ReadExistNodeDelay(file = i),File = KeyValFi)
        ExistDat['PM'].to_excel(writer,'ExistingPM__DelayRes')
writer.save()

VissimDATA = ReadExistNodeDelay(ExistingFile[0])
File = KeyValFi
subprocess.Popen([OutFi],shell=True)  