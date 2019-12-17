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

import os
import pandas as pd

#os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\12th-Street-TransitWay\Results')
os.chdir(r'H:\20\20548 - Arlington County Engineering On-Call\009 - 12th Street Transitway Extension\vissim\Results')

def CleanNetPerFun(file):
    '''
  
    '''
    Dat = pd.read_csv(file,sep =';',skiprows=1,comment="*")
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
    return(Dat)


#*********************************************************************************
# Specify Files
#*********************************************************************************

x1 = pd.ExcelFile(r'ResultsNameMap.xlsx')

NetPerfFile_AM = r'./RawVissimOutput/20548_2019_am-existing_V5_Vehicle Network Performance Evaluation Results.att'
file = NetPerfFile_AM
NetPerfFile_PM = NetPerfFile_AM



DataDict = {
        'ExistAM': NetPerfFile_AM,
        'ExistPM': NetPerfFile_PM,
        }


OutFi = r'NetworkPerfEvalRes.xlsx'

writer=pd.ExcelWriter(OutFi)
for name, file, in DataDict.items():
    Dat = CleanNetPerFun(file)
    Dat.to_excel(writer,name,na_rep=' ')
writer.save()    

        

