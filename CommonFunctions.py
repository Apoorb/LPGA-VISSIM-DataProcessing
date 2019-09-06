# -*- coding: utf-8 -*-
"""
Created on Fri Sep  6 13:03:48 2019

@author: abibeka
#Common Functions 
"""
import pandas as pd

# Read VISSIM File
#*********************************************************************************
def ReadExistNodeDelay(file):
    '''
    file: Node Evaluation file
    Returns: Cleaned resuls
    Read the Node-Evaluation File
    Ignore the ExistingPMfi. Can be used all AM and PM existing Node Evaluation
    '''
    ExistingPMDat=pd.read_csv(file,sep =';',skiprows=28)
    ExistingPMDat.columns
    #Use Avg values only
    ExistingPMDat = ExistingPMDat[(ExistingPMDat['$MOVEMENTEVALUATION:SIMRUN'] == 'AVG')]
    ExistingPMDat.rename(columns={'MOVEMENT':'Mvmt'},inplace=True)
    
    # Use the data for movements only. Ignore Intersection summary for now
    ExPMMvMDat = ExistingPMDat[ExistingPMDat.Mvmt.str.len()>1]
    
    ExPMMvMDat.loc[:,"HourInt"] = 'Nan'
    ExPMMvMDat.loc[:,'HourInt'] = ExPMMvMDat.TIMEINT
    ExPMMvMDat.HourInt = pd.Categorical(ExPMMvMDat.HourInt,['900-1800','1800-2700','2700-3600','3600-4500',
                                        '4500-5400','5400-6300','6300-7200','7200-8100',
                                        '8100-9000','9000-9900','9900-10800','10800-11700'])
    ExPMMvMDat = ExPMMvMDat.groupby(['HourInt','Mvmt']).aggregate({'VEHS(ALL)':'sum','QLENMAX':'max','VEHDELAY(ALL)':'first'}).reset_index()
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
    ExPMMvMDat.rename(columns= {'VEHS(ALL)':'veh','QLENMAX':'QLenMax','VEHDELAY(ALL)':'Delay'},inplace=True)
    return(ExPMMvMDat)
    
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
        
def ReadMergeVissimObs(VissimDATA, File):
    '''
    Don't care about ExistingPM_Vissim nomeclature
    VissimDATA = ExPMMvMDat,  # Vissim results to merge
    File = KeyValFi, # Key value file for PM
    Returns = Result reshaped for Report
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
    ExistPM_Vissim.HourInt = pd.Categorical(ExistPM_Vissim.HourInt,['900-1800','1800-2700','2700-3600','3600-4500',
                                        '4500-5400','5400-6300','6300-7200','7200-8100',
                                        '8100-9000','9000-9900','9900-10800','10800-11700'])
    ExistPM_Vissim = ExistPM_Vissim.sort_values(['Intersection','HourInt','Movement'])
    ExistPM_Vissim.loc[:,'DelayIntoVeh'] = ExistPM_Vissim.veh *ExistPM_Vissim.Delay
    ExistPM_Vissim = ExistPM_Vissim.groupby(['Intersection','HourInt','Movement']).aggregate({'QLenMax':'max','DelayIntoVeh':'sum','veh':'sum'}).reset_index()
    TempOverall = ExistPM_Vissim.groupby(['Intersection','HourInt']).aggregate({'QLenMax':'max','DelayIntoVeh':'sum','veh':'sum'}).reset_index()
    TempOverall.loc[:,'Movement'] = 'OverallIntersection'
    ExistPM_Vissim = pd.concat([ExistPM_Vissim,TempOverall],sort=False).reset_index()
    ExistPM_Vissim.loc[:,'Delay'] = (ExistPM_Vissim.DelayIntoVeh/ ExistPM_Vissim.veh).round(1).astype('float')
    ExistPM_Vissim.loc[:,'Delay'] = ExistPM_Vissim.loc[:,'Delay'].fillna(0)
    ExistPM_Vissim.loc[:,'LOS'] = ExistPM_Vissim.Delay.apply(LOS_Calc)
    ExistPM_Vissim.loc[:,'QLenMax'] = ExistPM_Vissim.QLenMax.round(1)
    ExistPM_Vissim.drop(columns=['veh','DelayIntoVeh'],inplace=True)
    ExistPM_Vissim.Movement = pd.Categorical(ExistPM_Vissim.Movement,[
            'EB U/L','EB L/T','EB T', 'EB R'
    ,'WB L','WB U/L', 'WB T/R','WB T', 'WB R','NB U/L', 'NB L','NB T/R',
     'NB T', 'NB R','SB U/L','SB L', 'SB T','SB R','SB T/R','SB L/T/R','OverallIntersection'
    ])
# Break Data in 1 hour Groups
    mask1 = ExistPM_Vissim.HourInt.isin(['900-1800','1800-2700','2700-3600','3600-4500'])
    mask2 = ExistPM_Vissim.HourInt.isin(['4500-5400','5400-6300','6300-7200','7200-8100'])
    mask3 = ExistPM_Vissim.HourInt.isin(['8100-9000','9000-9900','9900-10800','10800-11700'])
    ExistPM_Vissim.loc[:,'HourGroup'] = -99
    ExistPM_Vissim.loc[mask1,'HourGroup'] = 1
    ExistPM_Vissim.loc[mask2,'HourGroup'] = 2
    ExistPM_Vissim.loc[mask3,'HourGroup'] = 3
    #Rename the intersections
    ExistPM_Vissim.set_index('Intersection',inplace=True)
    #Reshaped the Data
    IntersectionKey = {'1':'LPGA/Tomoka Farms',
                    '2':'LPGA/I-95 SB Ramp',
                    '3':'LPGA/I-95 NB Ramp',
                    '4':'LPGA/Technology',
                    '5':'LPGA/Williamson',
                    '6':'LPGA/Clyde Morris'}
    ExistPM_Vissim =ExistPM_Vissim.rename(axis = 0, mapper=IntersectionKey).reset_index()
    ExistPM_Vissim.Intersection = pd.Categorical(ExistPM_Vissim.Intersection,[
                        'LPGA/Tomoka Farms',
                        'LPGA/I-95 SB Ramp',
                        'LPGA/I-95 NB Ramp',
                        'LPGA/Technology',
                        'LPGA/Williamson',
                        'LPGA/Clyde Morris'])
    
    ExistPM_Vissim = ExistPM_Vissim.set_index(['Intersection','Movement','HourInt','HourGroup']).unstack().unstack()
    ExistPM_Vissim.columns = ExistPM_Vissim.columns.swaplevel(0, 2)
    ExistPM_Vissim.columns = ExistPM_Vissim.columns.swaplevel(0, 1)
    mux = pd.MultiIndex.from_product([[1,2,3],['900-1800','1800-2700','2700-3600','3600-4500',
                                        '4500-5400','5400-6300','6300-7200','7200-8100',
                                        '8100-9000','9000-9900','9900-10800','10800-11700'],
                                  [u'Delay',u'LOS','QLenMax']], names=['HourGroup','HourInt',''])
    ExistPM_Vissim = ExistPM_Vissim.reindex(mux,axis=1)
    ExistPM_Vissim = ExistPM_Vissim.dropna(axis=1)
    return(ExistPM_Vissim)

    