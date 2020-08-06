# -*- coding: utf-8 -*-
"""
Created on Tue Dec 10 10:38:22 2019

@author: abibeka
"""


#0.0 Housekeeping. Clear variable space
from IPython import get_ipython  #run magic commands
ipython = get_ipython()
ipython.magic("reset -f")
ipython = get_ipython()



import os, sys
import pandas as pd
import numpy as np
import glob
import re

#os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\12th-Street-TransitWay\Results')
os.chdir(r'H:\20\20548 - Arlington County Engineering On-Call\009 - 12th Street Transitway Extension\vissim\Results')

owd = os.getcwd()
x1 = pd.ExcelFile(r'ResultsNameMap.xlsx')
x1.sheet_names
DatColMap = x1.parse('DatColMap')

os.chdir('CollectedData/149795 60 Minute Reports')
files = glob.glob('* [EWNS]B Speed.csv')


file = files[5]
temp =  files[0].split('-')
Location = temp[1].strip()
Dir = temp[2].split(" ")[1]

#******************************************************************************************************************************************************************
RawDataDict = {}
Data = pd.DataFrame(columns =['Dir','Location','Peak','AvgSpd'])
for file in files:
    temp =  file.split('-')
    Location = temp[1].strip()
    Dir = temp[2].split(" ")[1]

    Dat = pd.read_csv(file, skiprows= 14,nrows = 25)
    Dat = Dat.iloc[:,:-3]
    Dat.rename(columns = {"Date: ":'Col'},inplace=True)
    Dat.iloc[0,0] = "LowerLim"
    Dat.iloc[1,0] = "UpperLim"
    Dat = Dat[(Dat.Col == 'LowerLim') | (Dat.Col == 'UpperLim')|(Dat.Col == '08:00 AM') | (Dat.Col == '05:00 PM')]
    Dat = Dat.set_index('Col').T.reset_index(drop = True)
    Dat.loc[:,"AvgSpd"] = (Dat.UpperLim +Dat.LowerLim)/2
    RawDataDict[Location] = Dat
    AM_AvgSpd = sum(Dat['08:00 AM'] *Dat.AvgSpd)/ sum(Dat['08:00 AM'] )
    PM_AvgSpd = sum(Dat['05:00 PM'] *Dat.AvgSpd)/ sum(Dat['05:00 PM'] )
    Data = Data.append(pd.Series([Dir,Location,"AM Peak",AM_AvgSpd ],index=Data.columns), ignore_index=True)
    Data = Data.append(pd.Series([Dir,Location,"PM Peak",PM_AvgSpd ],index=Data.columns), ignore_index=True)
#******************************************************************************************************************************************************************


#Write Output
os.getcwd()
os.chdir('..')
os.getcwd()

outFi = 'ATR_AM_PM_Speeds.xlsx'
writer=pd.ExcelWriter(outFi)
for key, dat in RawDataDict.items():
    dat.to_excel(writer, key ,na_rep='-')
Data.to_excel(writer, "SAvgSpdField" ,na_rep='-')
writer.save() 











