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
ExistingFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Existing\*_Vehicle Network Performance Evaluation Results.att')
NoBuildFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\No Build\*_Vehicle Network Performance Evaluation Results.att')
DCDCFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCDI\*_Vehicle Network Performance Evaluation Results.att')
DCMIFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCMI\*_Vehicle Network Performance Evaluation Results.att')


ExistDat = {}
ExistDat['AM'] 

Dat = pd.read_csv(ExistingFile[0],sep =';',skiprows=62)
Dat.columns
Dat = Dat[Dat['$VEHICLENETWORKPERFORMANCEMEASUREMENTEVALUATION:SIMRUN']=='AVG']

Dat.rename(columns = {'TIMEINT':'TimeInt','DELAYAVG(ALL)': 'DelayAvg(All)','':''})