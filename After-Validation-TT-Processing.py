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
            15:'SB I-95 (SR-40 to SB Off-Ramp)',
            16:'SB I-95 (SB Off-Ramp to SB Loop-Ramp)',
            17:'SB I-95 (SB Loop-Ramp to SB On-Ramp)',
            18:'SB I-95 (SB On-Ramp to US-92)',
            19:'NB I-95 (US-92 to NB Off-Ramp)',
            20:'NB I-95 (NB Off-Ramp to NB Loop-Ramp)',
            21:'NB I-95 ( NB Loop-Ramp to NB On-Ramp)',
            22:'NB I-95 (NB On-Ramp to SR-40)'}
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
    TTData.loc[:,"VissimSMS"] = TTData.groupby(['TTSegNm'])['VissimSMS'].transform(lambda v: v.ffill())

    return(TTData)

# Get the Weighted Density by Segments
#Debug:
#file = NoBuildAMfi
#SegKeyVal = TTSegLaneDat
#SegIN = SegIn1
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
    ExistingAMDat.loc[:,'FlowRate_PerLane'] = ExistingAMDat.Veh *4/ ExistingAMDat.NumLanes
    ExistingAMDat.loc[:,"VissimSMS"] = (ExistingAMDat['Len']/ExistingAMDat['VissimTT']/1.47).round(1)
    ExistingAMDat.loc[:,'Density'] = (ExistingAMDat.FlowRate/ ExistingAMDat.VissimSMS).round(1)
    ExistingAMDat.loc[:,'Density_Len_Lanes'] = ExistingAMDat.Density *ExistingAMDat.Len
    ExistingAMDat.loc[:,'Len_Lanes'] = ExistingAMDat.Len * ExistingAMDat.NumLanes
    ExistingAMDat.columns
    DensityData = ExistingAMDat.groupby(['TIMEINT','SegName']).aggregate({'Len_Lanes':'sum','Density_Len_Lanes':'sum','VissimSMS':'mean','FlowRate_PerLane':'mean'}).reset_index()
    #DensityData = ExistingAMDat.groupby(['TIMEINT','SegName'])['Len_Lanes','Density_Len_Lanes'].sum().reset_index()
    DensityData.loc[:,'WeightedDensity'] = (DensityData.Density_Len_Lanes/ DensityData.Len_Lanes).round(1)
    DensityData.loc[:,"WeightedDensity"] = DensityData.groupby(['SegName'])['WeightedDensity'].transform(lambda v: v.ffill())
    DensityData.loc[:,"VissimSMS"] = DensityData.groupby(['SegName'])['VissimSMS'].transform(lambda v: v.ffill())
    DensityData = DensityData[['TIMEINT','SegName','WeightedDensity']]
    return(DensityData)


I95_Segs = [     'NB I-95 (US-92 to NB Off-Ramp)',
                'NB I-95 (NB Off-Ramp to NB Loop-Ramp)',
                'NB I-95 ( NB Loop-Ramp to NB On-Ramp)',
                'NB I-95 (NB On-Ramp to SR-40)',
                'SB I-95 (SR-40 to SB Off-Ramp)',
                'SB I-95 (SB Off-Ramp to SB Loop-Ramp)',
                'SB I-95 (SB Loop-Ramp to SB On-Ramp)',
                'SB I-95 (SB On-Ramp to US-92)'
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

# VISSIM File
#*********************************************************************************
PathToFile= r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2'
ExistingPMfi = r'Existing\20834_Existing_PM--C1C2C3C4C5C6C7C8_Vehicle Travel Time Results.att'
ExistingPMfi = os.path.join(PathToFile,ExistingPMfi)
ExistingAMfi =r'Existing\20834_Existing_AM--C1C2aC3C4C5C6C7C8C9_Vehicle Travel Time Results.att'
ExistingAMfi = os.path.join(PathToFile,ExistingAMfi)



NoBuildAMfi = r'No Build\20834_NoBuild_AM--C1C2C3C4C5C6C7C8C9_Vehicle Travel Time Results.att'
NoBuildPMfi = r'No Build\20834_NoBuild_PM---C1C2C3C4C5C6C7C8_Vehicle Travel Time Results.att'
BuildDCDI_AMfi = r'Build\DCDI\20834_Build1_6LnMod_AM---C1C2C3C4C5C6C7C8C9C10_Vehicle Travel Time Results.att'
BuildDCDI_PMfi = r'Build\DCDI\20834_Build1_6LnMod_PM---C1C2C3C4C5C6C7C8C9_Vehicle Travel Time Results.att'
BuildDCMI_AMfi = r'Build\DCMI\20834_DCMI_AM_20180808--C1C2C3C4C5C6C7C8_Vehicle Travel Time Results.att'
BuildDCMI_PMfi = r'Build\DCMI\20834_DCMI_PM_20180808--C1C2C3C4C5C6C7C8_Vehicle Travel Time Results.att'

NoBuildAMfi = os.path.join(PathToFile,NoBuildAMfi)
NoBuildPMfi = os.path.join(PathToFile,NoBuildPMfi)
BuildDCDI_AMfi = os.path.join(PathToFile,BuildDCDI_AMfi)
BuildDCDI_PMfi = os.path.join(PathToFile,BuildDCDI_PMfi)
BuildDCMI_AMfi = os.path.join(PathToFile,BuildDCMI_AMfi)
BuildDCMI_PMfi = os.path.join(PathToFile,BuildDCMI_PMfi)

#*******************************************************************************
#*******************************************************************************
# Which File are we analysising. Only make changes here and done:
FileAM = ExistingAMfi
FileAM  = NoBuildAMfi
FileAM  = BuildDCDI_AMfi
FileAM  = BuildDCMI_AMfi
#File Being Processed
#*******************************************************************************
FileAM = ExistingAMfi
#*******************************************************************************

def switch_PMFile(argument):
    switcher = {
        ExistingAMfi: ExistingPMfi,
        NoBuildAMfi: NoBuildPMfi,
        BuildDCDI_AMfi: BuildDCDI_PMfi,
        BuildDCMI_AMfi: BuildDCMI_PMfi
    }
    return switcher.get(argument, "Invalid Name")

FilePM = switch_PMFile(FileAM)

def switch_OutputTag(argument):
    switcher = {
        ExistingAMfi: "Existing",
        NoBuildAMfi: "No-Build",
        BuildDCDI_AMfi: "DCDI",
        BuildDCMI_AMfi: "DCMI"
    }
    return switcher.get(argument, "Invalid Name")
    
OutputFolderName = switch_OutputTag(FileAM)

if((FileAM == BuildDCDI_AMfi) | (FileAM == BuildDCMI_AMfi)):
    AreResultsForBuild = True
else: AreResultsForBuild = False
# Get the time keys to convert Vissim intervals to Hours of the day
PathToVis = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
TimeConvFi  = 'TravelTimeKeyValuePairs.xlsx'
TimeConvFi = pd.ExcelFile(os.path.join(PathToVis,TimeConvFi))
TimeConvFi.sheet_names
TimeKeys = TimeConvFi.parse('TimeVissimKey')

if(AreResultsForBuild):
    TTSegLaneDat = TimeConvFi.parse('Build')
    TTSegLaneDat = TTSegLaneDat.dropna()
else:
    # Get The Travel time (TT) segment and # of lane data
    TTSegLaneDat = TimeConvFi.parse('Existing')
    TTSegLaneDat = TTSegLaneDat.dropna()


#Summarize the Travel time for AM and PM. Concatenate the 2 datasets.
SegIn1 = I95_Segs+ LPGASeg
TTAM = PreProcessVissimTT(FileAM,SegIn1)
TTPM = PreProcessVissimTT(FilePM,SegIn1)
DenAM = PreProcessVissimDensity(FileAM,TTSegLaneDat)
DenPM = PreProcessVissimDensity(FilePM,TTSegLaneDat)
TTAM = TTAM.merge(DenAM,left_on=['TIMEINT','TTSegNm'],right_on=['TIMEINT','SegName'],how='left')
TTPM = TTPM.merge(DenPM,left_on=['TIMEINT','TTSegNm'],right_on=['TIMEINT','SegName'],how='left')
TTAM.loc[:,'AnalysisPeriod'] = 'AM'
TTPM.loc[:,'AnalysisPeriod'] = 'PM'
Vissim_TT = pd.concat([TTAM,TTPM],sort=False)
Vissim_TT.columns
Vissim_TT = Vissim_TT[['AnalysisPeriod','TIMEINT','TTSegNm','VissimSMS','WeightedDensity','Veh','Len']]

#Correction-Oversaturated-Conditions
# For 10 to 20 mph if the density > 45, just use it else replace with 160
mask = (Vissim_TT.VissimSMS < 20) & (Vissim_TT.WeightedDensity < 45)
Vissim_TT.loc[mask,"WeightedDensity"] = 160 # 33 ft. spacing
# For < 10 mph replace all values with 160
mask = (Vissim_TT.VissimSMS < 10) & (Vissim_TT.WeightedDensity < 200)
Vissim_TT.loc[mask,"WeightedDensity"] = 160 # 33 ft. spacing
Debug = Vissim_TT.loc[mask,"WeightedDensity"]

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
PerfMes ='Density'
(not((d=='EB' or d=='WB') & (PerfMes=='Density')))
if not os.path.exists(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\Figs\{}'.format(OutputFolderName)):
    os.makedirs(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\Figs\{}'.format(OutputFolderName))
os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\Figs\{}'.format(OutputFolderName))
os.getcwd()
DictFigs = {}
idx = pd.IndexSlice
#Color Palettes
#https://matplotlib.org/users/colormaps.html
for t in TimePer:
    for d in Direction:
        for PerfMes in PerfMeasure:
            if(not((d=='EB' or d=='WB') & (PerfMes=='Density'))):
                HeatMapDat = Vissim_TT.loc[idx[:,d,:], idx[t,['Time',PerfMes]]]
                HeatMapDat = HeatMapDat.droplevel(0,axis=1).reset_index()
                HeatMapDat = HeatMapDat.pivot('Time','TTSegNm',PerfMes)
                HeatMapDat = HeatMapDat.astype('int32')
                plt.figure()
                if PerfMes=='Density':
                    title_ = "Density (veh/mile/lane)"
                    colorBar_ = 'viridis_r'
                    vmax_ = 160; vmin_ = 0
                    g = sns.heatmap(HeatMapDat,cmap = colorBar_, linewidths=.5,vmin=vmin_, vmax=vmax_, annot=True,fmt="d")
                else:   
                    title_ = "Space Mean Speed (mph)"
                    colorBar_ = 'viridis'
                    if(d=='SB') | (d=='NB'):
                        vmax_ = 70; vmin_ = 0
                    else:
                        vmax_ = 45; vmin_ = 0
                    g = sns.heatmap(HeatMapDat,cmap = colorBar_, linewidths=.5, vmin=vmin_, vmax=vmax_, annot=True,fmt="d")
                
                # figure size in inches
                #sns.set(rc={'figure.figsize':(6,10)})
                
                g.set_xticklabels(rotation=30,labels = g.get_xticklabels(),ha='right')
                g.set_title('{}—{}'.format(t,title_))
                g.set_xlabel("")
                g.set_ylabel("Time Interval")
                fig = g.get_figure()
                fig.savefig("".join([t,d,PerfMes,'.jpg']),bbox_inches="tight")


if not os.path.exists(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\Figs\{}'.format(OutputFolderName)):
    os.makedirs(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\Figs\{}'.format(OutputFolderName))
os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\Figs\{}'.format(OutputFolderName))
os.getcwd()
t = 'AM'
d = 'SB'
PerfMes ='VissimSMS'
for t in TimePer:
    for d in Direction:
        for PerfMes in PerfMeasure:
            os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\Figs\{}'.format(OutputFolderName))
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
                        ax1.set_ylim(0,170)
                        ylab = "Density (veh/mile/lane)"

                    ax1.set_title('{}—{}'.format(t,Tseg), y =1.05)
                    ax1.set_ylabel(ylab)
                    ax1.set_xlabel("Time Interval")
                    # Must draw the canvas to position the ticks
                    fig.canvas.draw()
                    ax1.set_xticklabels(rotation=30,labels = ax1.get_xticklabels(),ha='right')
                    fig.savefig("_".join([Tseg,d,PerfMes,'.jpg']),bbox_inches="tight")
                    plt.close()
#Write Results to Output File
PathToOutFi = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files'
OutFi = "AfterValidationRes.xlsx"
OutFi = os.path.join(PathToOutFi,OutFi)
if(OutputFolderName =='Existing'):
    writer=pd.ExcelWriter(OutFi)
else:
    writer=pd.ExcelWriter(OutFi, mode ='a')
Vissim_LPGA.to_excel(writer, '{}-LPGA-SMS-Res'.format(OutputFolderName),na_rep='NaN')
Vissim_I95.to_excel(writer,'{}-I95-SMS&Density-Res'.format(OutputFolderName),na_rep='NaN')
writer.save()

if(OutputFolderName == 'DCMI'):
    subprocess.Popen([OutFi],shell=True)  

