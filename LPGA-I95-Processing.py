# -*- coding: utf-8 -*-
"""
Created on Tue Aug 27 08:04:07 2019

@author: abibeka
# Get Travel time and TMC on I-95
"""

#0.0 Housekeeping. Clear variable space
from IPython import get_ipython  #run magic commands
ipython = get_ipython()
ipython.magic("reset -f")
ipython = get_ipython()



import os
import pandas as pd
import numpy as np
import subprocess

# VISSIM File
#*********************************************************************************
PathToExist = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Existing'
ExistingPMfi = '20834_Existing_PM--C1C2C3C4C5C6C7_Vehicle Travel Time Results.att'
ExistingPMfi = os.path.join(PathToExist,ExistingPMfi)
ExistingAMfi ='20834_Existing_AM--C1C2aC3C4C5C6C7C8_Vehicle Travel Time Results.att'
ExistingAMfi = os.path.join(PathToExist,ExistingAMfi)



def TTSegName(x):
    TTSeg= {1: 'EB',
            2: 'WB',
            3: 'EB LPGA (Tomoka Rd to I-95 SB Ramp)',
            4: 'EB LPGA (I-95 SB Ramp to I-95 NB Ramp)',
            5: 'EB LPGA (I-95 NB Ramp to Technology Blvd)',
            6: 'EB LPGA (Technology Blvd to Willamson Blvd)',
            7: 'EB LPGA (Willamson Blvd to Clyde-Morris Blvd)',
            8: 'WB LPGA (Clyde-Morris Blvd to Willamson Blvd)',
            9: 'WB LPGA (Willamson Blvd to Technology Blvd)',
            10:'WB LPGA (Technology Blvd to I-95 NB Ramp)',
            11:'WB LPGA (I-95 NB Ramp to I-95 SB Ramp)',
            12:'WB LPGA (I-95 SB Ramp to Tomoka Rd)',
            13:'SB I-95',
            14:'NB I-95',
            15:'SB I-95 (SR40 to SB OffRamp)',
            16:'SB I-95 (SB OffRamp to SB LoopRamp)',
            17:'SB I-95 (SB LoopRamp to SB On-Ramp)',
            18:'SB I-95 (SB On-Ramp to US92)',
            19:'NB I-95 (US92 to NB OffRamp)',
            20:'NB I-95 (NB OffRamp to NB LoopRamp)',
            21:'NB I-95 ( NB LoopRamp to NB On-Ramp)',
            22:'NB I-95 (NB On-Ramp to SR40)'}
    Nm = TTSeg[x]
    return Nm




def PreProcessVissimTT(file):
    ExistingAMDat=pd.read_csv(file,sep =';',skiprows=17)
    ExistingAMDat.columns
    ExistingAMDat.rename(columns={'TRAVTM(ALL)':'VissimTT','VEHS(ALL)':'Veh'},inplace=True)
    mask=ExistingAMDat["$VEHICLETRAVELTIMEMEASUREMENTEVALUATION:SIMRUN"]=="AVG"
    ExistingAMDat = ExistingAMDat[mask]
    ExistingAMDat["TTSegNm"]=ExistingAMDat['VEHICLETRAVELTIMEMEASUREMENT'].apply(TTSegName)
    I95_Segs = ['SB I-95',
                'NB I-95',
                'SB I-95 (SR40 to SB OffRamp)',
                'SB I-95 (SB OffRamp to SB LoopRamp)',
                'SB I-95 (SB LoopRamp to SB On-Ramp)',
                'SB I-95 (SB On-Ramp to US92)',
                'NB I-95 (US92 to NB OffRamp)',
                'NB I-95 (NB OffRamp to NB LoopRamp)',
                'NB I-95 ( NB LoopRamp to NB On-Ramp)',
                'NB I-95 (NB On-Ramp to SR40)']
    ExistingAMDat.TIMEINT = pd.Categorical(ExistingAMDat.TIMEINT,['900-1800','1800-2700','2700-3600','3600-4500',
                                        '4500-5400','5400-6300','6300-7200','7200-8100',
                                        '8100-9000','9000-9900','9900-10800','10800-11700'])
    
    mask = ExistingAMDat.TTSegNm.isin(I95_Segs)
    ExistingAMDat = ExistingAMDat[mask]
    ExistingAMDat.loc[:,"VissimSMS"] = (ExistingAMDat['DISTTRAV(ALL)']/ExistingAMDat['VissimTT']/1.47).round(1)
    mask1 = ExistingAMDat['VEHICLETRAVELTIMEMEASUREMENT'].isin([13,14])
    TTData = ExistingAMDat[mask1]
    TMCData = ExistingAMDat[~mask1]
    return(TTData,TMCData)


file = ExistingAMfi
TTAM, TMCAM = PreProcessVissimTT(ExistingAMfi)
TTPM, TMCPM = PreProcessVissimTT(ExistingPMfi)

TTAM.loc[:,'AnalysisPeriod'] = 'AM'
TTPM.loc[:,'AnalysisPeriod'] = 'PM'
TMCAM.loc[:,'AnalysisPeriod'] = 'AM'
TMCPM.loc[:,'AnalysisPeriod'] = 'PM'


Vissim_TT = pd.concat([TTAM,TTPM],sort=False)
Vissim_TT.columns
Vissim_TT = Vissim_TT[['AnalysisPeriod','TIMEINT','TTSegNm','VissimSMS','Veh']]


#Read NPMRDS TT data
#*********************************************************************************
PathToNpmrds = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\I-95-TravelTime'
SpeedDat = pd.read_csv(os.path.join(PathToNpmrds,'I-95-TravelTime.csv'),usecols=['tmc_code','measurement_tstamp','travel_time_minutes'])
SpeedKey = pd.read_csv(os.path.join(PathToNpmrds,'TMC_Identification.csv'),usecols=['tmc', 'road', 'direction', 'miles'])
SpeedDat = SpeedDat.merge(SpeedKey, left_on='tmc_code',right_on='tmc',how='left')

SpeedDat.columns
SpeedDat.dtypes
SpeedDat.set_index(pd.DatetimeIndex(SpeedDat['measurement_tstamp']),inplace=True)

SpeedDat = SpeedDat.loc['2017-10-5'] # This is the date when ZHB did travel time runs
AMPeriod = ('6:15','9:00')
PMPeriod = ('15:30','18:15')

SpeedDat = pd.concat([SpeedDat.between_time(AMPeriod[0],AMPeriod[1]),SpeedDat.between_time(PMPeriod[0],PMPeriod[1])])
SpeedDat["AnalysisPeriod"] = np.nan
SpeedDat.loc[SpeedDat.between_time(AMPeriod[0],AMPeriod[1]).index,"AnalysisPeriod"] = "AM"
SpeedDat.loc[SpeedDat.between_time(PMPeriod[0],PMPeriod[1]).index,"AnalysisPeriod"] = "PM"

SpeedDat.reset_index(drop=True,inplace=True)
SpeedDat = SpeedDat.groupby(["AnalysisPeriod",'measurement_tstamp','direction'])['travel_time_minutes','miles'].sum().reset_index()
SpeedDat.loc[:,"NpmrdsSMS"] = (60*SpeedDat.miles/SpeedDat.travel_time_minutes).round(1)

PathTorouting = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\Routing_LPGA_Clean_AXB-V3.xlsx'
RoutingWB = pd.ExcelFile(PathTorouting)
RoutingWB.sheet_names
TimeVissimKey = RoutingWB.parse('TimeVissimKey')
TimeVissimKey.set_index('Time',inplace=True)
SpeedDat['time'] = pd.to_datetime(SpeedDat['measurement_tstamp']).dt.time
SpeedDat['TIMEINT'] =  TimeVissimKey.loc[SpeedDat['time']].values
SpeedDat.loc[SpeedDat.direction =='NORTHBOUND','direction']  = 'NB I-95'
SpeedDat.loc[SpeedDat.direction =='SOUTHBOUND','direction']  = 'SB I-95'
SpeedDat.rename(columns={'direction':'TTSegNm'},inplace=True)
SpeedDat = SpeedDat[['AnalysisPeriod','time','TIMEINT','TTSegNm','NpmrdsSMS']]
SpeedDat.columns
lkeys = ['AnalysisPeriod','TIMEINT','TTSegNm']

#Merge Data
#*********************************************************************************
SpeedDat = SpeedDat.merge(Vissim_TT,left_on=lkeys,right_on=lkeys,how = 'inner')
SpeedDat.loc[:,'DiffSMS'] =  SpeedDat.VissimSMS - SpeedDat.NpmrdsSMS
SpeedDat = pd.pivot_table(SpeedDat,index='TIMEINT',columns= ['AnalysisPeriod','TTSegNm'],values=['VissimSMS','NpmrdsSMS','DiffSMS'])


SpeedDat.columns = SpeedDat.columns.swaplevel(0,2)
SpeedDat.columns = SpeedDat.columns.swaplevel(0,1)
mux = pd.MultiIndex.from_product([['AM','PM'],['NB I-95','SB I-95'],
        ['NpmrdsSMS','VissimSMS','DiffSMS']], names=SpeedDat.columns.names)
SpeedDat = SpeedDat.reindex(mux,axis=1)
SpeedDat.index = pd.Categorical(SpeedDat.index,['900-1800','1800-2700','2700-3600','3600-4500',
                                        '4500-5400','5400-6300','6300-7200','7200-8100',
                                        '8100-9000','9000-9900','9900-10800','10800-11700'])
SpeedDat.sort_index(inplace=True)


#TMC Data
#*********************************************************************************
ObsVolumes = RoutingWB.parse('I95Vols',index_col=0,header =[0,1,2])
idx = pd.IndexSlice
ObsVolumes = ObsVolumes.loc[:,idx['Existing',:,:]]
ObsVolumes = ObsVolumes.unstack().reset_index().drop(columns='Year')
ObsVolumes.rename(columns = {0:'ObsFlowRate'},inplace=True)
Vissim_TMC = Vissim_TT[['AnalysisPeriod','TIMEINT','TTSegNm','Veh']]
Vissim_TMC.loc[:,'VissimFlowRate'] = Vissim_TMC.Veh*4
Vissim_TMC.drop(columns='Veh',inplace=True)
lkeys = ['AnalysisPeriod','TIMEINT','TTSegNm']
ObsVolumes = ObsVolumes.merge(Vissim_TMC,left_on=lkeys,right_on=lkeys,how = 'inner')
ObsVolumes.loc[:,'DiffFlowRate'] =  ObsVolumes.VissimFlowRate - ObsVolumes.ObsFlowRate
ObsVolumes = pd.pivot_table(ObsVolumes,index='TIMEINT',
                          columns= ['AnalysisPeriod','TTSegNm'],values=['ObsFlowRate','VissimFlowRate','DiffFlowRate'])
ObsVolumes.columns = ObsVolumes.columns.swaplevel(0,2)
ObsVolumes.columns = ObsVolumes.columns.swaplevel(0,1)
mux = pd.MultiIndex.from_product([['AM','PM'],['NB I-95','SB I-95'],
        ['ObsFlowRate','VissimFlowRate','DiffFlowRate']], names=ObsVolumes.columns.names)
ObsVolumes = ObsVolumes.reindex(mux,axis=1)
ObsVolumes.index = pd.Categorical(ObsVolumes.index,['900-1800','1800-2700','2700-3600','3600-4500',
                                        '4500-5400','5400-6300','6300-7200','7200-8100',
                                        '8100-9000','9000-9900','9900-10800','10800-11700'])
ObsVolumes.sort_index(inplace=True)
ObsVolumes


#Write to the Results File
#*********************************************************************************
PathToKeyVal = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
OutFi = "Report-TT-GEH-Results.xlsx"
OutFi = os.path.join(PathToKeyVal,OutFi)
writer=pd.ExcelWriter(OutFi,mode ='a')
SpeedDat.to_excel(writer, 'I95_TravelTime')
ObsVolumes.to_excel(writer, 'I95_FlowRate')
writer.save()

subprocess.Popen([OutFi],shell=True)  