# -*- coding: utf-8 -*-
"""
Created on Tue Aug 27 08:04:07 2019

@author: abibeka
# Get Travel time on LPGA and I-95 (After Validation)
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
import seaborn as sns
import matplotlib.pyplot as plt
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
    if x < 23:
        Nm = TTSeg[x]
    else:
        Nm = None
    return Nm




def PreProcessVissimTT(file, SegIN):
    '''
    file : VISSIM Results file
    SegIN : Specify segments for which results are needed
    Summarize Vissim Travel time results
    '''
    ExistingAMDat=pd.read_csv(file,sep =';',skiprows=17)
    ExistingAMDat.columns
    ExistingAMDat.rename(columns={'TRAVTM(ALL)':'VissimTT','VEHS(ALL)':'Veh','DISTTRAV(ALL)':'Len'},inplace=True)
    mask=ExistingAMDat["$VEHICLETRAVELTIMEMEASUREMENTEVALUATION:SIMRUN"]=="AVG"
    ExistingAMDat = ExistingAMDat[mask]
    ExistingAMDat["TTSegNm"]=ExistingAMDat['VEHICLETRAVELTIMEMEASUREMENT'].apply(TTSegName)
    ExistingAMDat.TIMEINT = pd.Categorical(ExistingAMDat.TIMEINT,['900-1800','1800-2700','2700-3600','3600-4500',
                                        '4500-5400','5400-6300','6300-7200','7200-8100',
                                        '8100-9000','9000-9900','9900-10800','10800-11700'])
    
    mask = ExistingAMDat.TTSegNm.isin(SegIN)
    ExistingAMDat = ExistingAMDat[mask]
    ExistingAMDat.loc[:,"VissimSMS"] = (ExistingAMDat['Len']/ExistingAMDat['VissimTT']/1.47).round(1)
    mask1 = ~ExistingAMDat['VEHICLETRAVELTIMEMEASUREMENT'].isin([13,14])
    TTData = ExistingAMDat[mask1]
    return(TTData)

# Get the Weighted Density by Segments
def PreProcessVissimDensity(file, SegKeyVal):
    '''
    file : VISSIM Results file
    SegKeyVal : Key value pair for segment # and TT segment name
    Summarize Vissim Density results
    '''
    ExistingAMDat=pd.read_csv(file,sep =';',skiprows=17)
    ExistingAMDat.columns
    ExistingAMDat.rename(columns={'TRAVTM(ALL)':'VissimTT','VEHS(ALL)':'Veh','DISTTRAV(ALL)':'Len'},inplace=True)
    mask=ExistingAMDat["$VEHICLETRAVELTIMEMEASUREMENTEVALUATION:SIMRUN"]=="AVG"
    ExistingAMDat = ExistingAMDat[mask]
    ExistingAMDat = SegKeyVal.merge(ExistingAMDat,left_on=['SegNO'],right_on = ['VEHICLETRAVELTIMEMEASUREMENT'],how= 'left')
    ExistingAMDat.TIMEINT = pd.Categorical(ExistingAMDat.TIMEINT,['900-1800','1800-2700','2700-3600','3600-4500',
                                        '4500-5400','5400-6300','6300-7200','7200-8100',
                                        '8100-9000','9000-9900','9900-10800','10800-11700'])
    ExistingAMDat.loc[:,"VissimSMS"] = (ExistingAMDat['Len']/ExistingAMDat['VissimTT']/1.47).round(1)
    #Get flow rate and density
    ExistingAMDat.loc[:,'FlowRate'] = ExistingAMDat.Veh *4
    ExistingAMDat.loc[:,'DensityPerLane'] = (ExistingAMDat.FlowRate/ ExistingAMDat.VissimSMS/ExistingAMDat.NumLanes).round(1)
    ExistingAMDat.loc[:,'LenByDensity'] = ExistingAMDat.DensityPerLane *ExistingAMDat.Len
    ExistingAMDat.columns
    DensityData = ExistingAMDat.groupby(['TIMEINT','SegName'])['Len','LenByDensity'].sum().reset_index()
    DensityData.loc[:,'WeightedDensity'] = (DensityData.LenByDensity/ DensityData.Len).round(1)
    DensityData.rename(columns={'Len':'Len1'},inplace=True)
    DensityData = DensityData[['TIMEINT','SegName','Len1','WeightedDensity']]
    return(DensityData)


I95_Segs = [     'NB I-95 (US92 to NB OffRamp)',
                'NB I-95 (NB OffRamp to NB LoopRamp)',
                'NB I-95 ( NB LoopRamp to NB On-Ramp)',
                'NB I-95 (NB On-Ramp to SR40)',
                'SB I-95 (SR40 to SB OffRamp)',
                'SB I-95 (SB OffRamp to SB LoopRamp)',
                'SB I-95 (SB LoopRamp to SB On-Ramp)',
                'SB I-95 (SB On-Ramp to US92)'
               ]
  
LPGASeg=['EB LPGA (Tomoka Rd to I-95 SB Ramp)',
    'EB LPGA (I-95 SB Ramp to I-95 NB Ramp)',
    'EB LPGA (I-95 NB Ramp to Technology Blvd)', 
    'EB LPGA (Technology Blvd to Willamson Blvd)',
    'EB LPGA (Willamson Blvd to Clyde-Morris Blvd)',
    'WB LPGA (Clyde-Morris Blvd to Willamson Blvd)', 
    'WB LPGA (Willamson Blvd to Technology Blvd)', 
    'WB LPGA (Technology Blvd to I-95 NB Ramp)',
    'WB LPGA (I-95 NB Ramp to I-95 SB Ramp)',
    'WB LPGA (I-95 SB Ramp to Tomoka Rd)']


# Get the time keys to convert Vissim intervals to Hours of the day
PathToVis = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
TimeConvFi  = 'TravelTimeKeyValuePairs.xlsx'
TimeConvFi = pd.ExcelFile(os.path.join(PathToVis,TimeConvFi))
TimeConvFi.sheet_names
TimeKeys = TimeConvFi.parse('TimeVissimKey')

# Get The Travel time (TT) segment and # of lane data
TTSegLaneDat = TimeConvFi.parse('Existing')
TTSegLaneDat = TTSegLaneDat.dropna()

#Summarize the Travel time for AM and PM. Concatenate the 2 datasets.
SegIn1 = I95_Segs+ LPGASeg
file = ExistingAMfi
TTAM = PreProcessVissimTT(ExistingAMfi,SegIn1)
TTPM = PreProcessVissimTT(ExistingPMfi,SegIn1)
DenAM = PreProcessVissimDensity(ExistingAMfi,TTSegLaneDat)
DenPM = PreProcessVissimDensity(ExistingPMfi,TTSegLaneDat)
TTAM = TTAM.merge(DenAM,left_on=['TIMEINT','TTSegNm'],right_on=['TIMEINT','SegName'],how='left')
TTPM = TTPM.merge(DenPM,left_on=['TIMEINT','TTSegNm'],right_on=['TIMEINT','SegName'],how='left')
TTAM.loc[:,'AnalysisPeriod'] = 'AM'
TTPM.loc[:,'AnalysisPeriod'] = 'PM'
Vissim_TT = pd.concat([TTAM,TTPM],sort=False)
Vissim_TT.columns
Vissim_TT = Vissim_TT[['AnalysisPeriod','TIMEINT','TTSegNm','VissimSMS','WeightedDensity','Veh','Len']]

# Convert Vissim intervals to Hours of the day

lkey_ = ['AnalysisPeriod','TIMEINT']
Vissim_TT = Vissim_TT.merge(TimeKeys,left_on = lkey_, right_on =lkey_,how = 'inner')


#Sort the Travel Segments in Correct Order
Vissim_TT.TTSegNm = pd.Categorical(Vissim_TT.TTSegNm,SegIn1)
Vissim_TT.Time = pd.Categorical(Vissim_TT.Time,TimeKeys.Time)
Vissim_TT.TIMEINT = pd.Categorical(Vissim_TT.TIMEINT,['900-1800','1800-2700','2700-3600','3600-4500',
                                        '4500-5400','5400-6300','6300-7200','7200-8100',
                                        '8100-9000','9000-9900','9900-10800','10800-11700'])
    
Vissim_TT.loc[:,'Dir'] = Vissim_TT.TTSegNm.apply(lambda x: x.split(' ')[0])
Vissim_TT.rename(columns={'WeightedDensity':'Density'},inplace=True)

Vissim_TT.sort_values(['AnalysisPeriod','TIMEINT','Dir','TTSegNm'],inplace=True)
Vissim_TT.set_index(['TIMEINT','Dir','TTSegNm','AnalysisPeriod'],inplace=True)
Vissim_TT = Vissim_TT.unstack()
Vissim_TT = Vissim_TT.swaplevel(0,1,axis=1)

mux = pd.MultiIndex.from_product([['AM','PM'],['Time','VissimSMS','Density','Veh','Len','FlowRate'],
        ], names=Vissim_TT.columns.names)
Vissim_TT = Vissim_TT.reindex(mux,axis=1)
idx = pd.IndexSlice
Vissim_I95 = Vissim_TT.loc[idx[:,:,I95_Segs], idx[:,['VissimSMS','Density']]] 
Vissim_LPGA =Vissim_TT.loc[idx[:,:,LPGASeg], idx[:,['VissimSMS']]]



TimePer = ['AM','PM']
Direction = ['NB','SB','EB','WB']
PerfMeasure = ['VissimSMS','Density']
t = 'AM'
d = 'SB'
PerfMes ='VissimSMS'
(not((d=='EB' or d=='WB') & (PerfMes=='Density')))
os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\Figs')
DictFigs = {}
idx = pd.IndexSlice
for t in TimePer:
    for d in Direction:
        for PerfMes in PerfMeasure:
            if(not((d=='EB' or d=='WB') & (PerfMes=='Density'))):
                HeatMapDat = Vissim_TT.loc[idx[:,d,:], idx[t,['Time',PerfMes]]]
                HeatMapDat = HeatMapDat.droplevel(0,axis=1).reset_index()
                HeatMapDat = HeatMapDat.pivot('Time','TTSegNm',PerfMes)
                if PerfMes=='Density':
                    vmax_ = 65; vmin_ = 0
                    title_ = "Density (veh/mile/lane)"
                else:   
                    title_ = "Space Mean Speed (mph)"
                    if(d=='SB') | (d=='NB'):
                        vmax_ = 70; vmin_ = 0
                    else:
                        vmax_ = 45; vmin_ = 0
                plt.figure()
                # figure size in inches
                #sns.set(rc={'figure.figsize':(6,10)})
                g = sns.heatmap(HeatMapDat,cmap = 'Greys', linewidths=.5, vmin=vmin_, vmax=vmax_)
                g.set_xticklabels(rotation=30,labels = g.get_xticklabels(),ha='right')
                g.set_title('{}'.format(title_))
                g.set_xlabel("")
                g.set_ylabel("15 min Time Interval")
                fig = g.get_figure()
                fig.savefig("".join([t,d,PerfMes,'.jpg']),bbox_inches="tight")

os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\Figs')
os.getcwd()
t = 'AM'
d = 'SB'
PerfMes ='VissimSMS'
for t in TimePer:
    for d in Direction:
        for PerfMes in PerfMeasure:
            os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\Figs')
            if not os.path.exists('{}/{}/{}'.format(t,d,PerfMes)):
                os.makedirs('{}/{}/{}'.format(t,d,PerfMes))
            os.chdir('{}/{}/{}'.format(t,d,PerfMes))
            if(not((d=='EB' or d=='WB') & (PerfMes=='Density'))):
                LinePlotDat = Vissim_TT.loc[idx[:,d,:], idx[t,['Time',PerfMes]]]
                LinePlotDat = LinePlotDat.droplevel(0,axis=1).reset_index()
                for Tseg in LinePlotDat.TTSegNm.unique():
                    TempDat = LinePlotDat[LinePlotDat.TTSegNm==Tseg]
                    sns.set(rc={'figure.figsize':(6,4)})
                    sns.set_style("ticks")
                    fig,ax1 = plt.subplots()
                    sns.lineplot('Time',PerfMes,data = TempDat,ax=ax1)
                    if(PerfMes == 'VissimSMS'):
                        ax1.set_ylim(0,70)
                        ylab = "Space Mean Speed (mph)"
                    else: 
                        ax1.set_ylim(0,65)
                        ylab = "Density (veh/mile/lane)"

                    ax1.set_title('{}'.format(Tseg), y =1.05)
                    ax1.set_ylabel(ylab)
                    ax1.set_xlabel("15 min Time Interval")
                    # Must draw the canvas to position the ticks
                    fig.canvas.draw()
                    ax1.set_xticklabels(rotation=30,labels = ax1.get_xticklabels(),ha='right')
                    fig.savefig("_".join([Tseg,d,PerfMes,'.jpg']),bbox_inches="tight")
                    plt.close()
#Write Results to Output File
PathToOutFi = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
OutFi = "AfterValidationRes.xlsx"
OutFi = os.path.join(PathToOutFi,OutFi)
writer=pd.ExcelWriter(OutFi)
Vissim_LPGA.to_excel(writer, 'LPGA-SMS-Res')
Vissim_I95.to_excel(writer,'I95-SMS&Density-Res')
writer.save()


subprocess.Popen([OutFi],shell=True)  
