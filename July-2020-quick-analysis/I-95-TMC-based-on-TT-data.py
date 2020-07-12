# -*- coding: utf-8 -*-
"""
Created on Mon Aug 19 13:47:01 2019

@author: abibeka
Read data from travel time files
"""

#0.0 Housekeeping. Clear variable space
# from IPython import get_ipython  #run magic commands
# ipython = get_ipython()
# ipython.magic("reset -f")
# ipython = get_ipython()



import os, numpy as np
import pandas as pd
from io import StringIO

#Observed data
ObsDataAmIO = StringIO("""TTSegNm;4500-5400;5400-6300;6300-7200;7200-8100
SB I-95 (SR40 to SB OffRamp);4357;5421;5421;4163
SB I-95 (SB OffRamp to SB LoopRamp);3176;3952;3952;3035
SB I-95 (SB LoopRamp to SB On-Ramp);3452;4261;4299;3333
SB I-95 (SB On-Ramp to US92);3511;4327;4372;3396
NB I-95 (US92 to NB OffRamp);3857;4801;4801;3686
NB I-95 (NB OffRamp to NB LoopRamp);3058;3806;3806;2922
NB I-95 ( NB LoopRamp to NB On-Ramp);3303;4080;4113;3186
NB I-95 (NB On-Ramp to SR40);3599;4412;4485;3506
    """)
ObsDataAm = pd.read_csv(ObsDataAmIO, sep =";")
ObsDataAm = ObsDataAm.set_index('TTSegNm').stack().reset_index()
ObsDataAm.columns = ['TTSegNm','HourInt','ObsVol']
ObsDataAm.ObsVol = ObsDataAm.ObsVol/4
ObsDataPmIO = StringIO("""TTSegNm;4500-5400;5400-6300;6300-7200;7200-8100
SB I-95 (SR40 to SB OffRamp);4108;3982;4401;4275
SB I-95 (SB OffRamp to SB LoopRamp);3408;3304;3651;3547
SB I-95 (SB LoopRamp to SB On-Ramp);3990;3956;4381;4175
SB I-95 (SB On-Ramp to US92);4038;4010;4441;4227
NB I-95 (US92 to NB OffRamp);3432;3328;3678;3573
NB I-95 (NB OffRamp to NB LoopRamp);2901;2813;3109;3020
NB I-95 ( NB LoopRamp to NB On-Ramp);3170;3114;3447;3310
NB I-95 (NB On-Ramp to SR40);3929;3964;4398;4129""")
ObsDataPm = pd.read_csv(ObsDataPmIO, sep =";")
ObsDataPm = ObsDataPm.set_index('TTSegNm').stack().reset_index()
ObsDataPm.columns = ['TTSegNm','HourInt','ObsVol']
ObsDataPm.ObsVol = ObsDataPm.ObsVol/4

# Use consistent naming in VISSIM
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
    if(x<23):
        Nm = TTSeg[x]
    else: Nm = None
    return Nm

# VISSIM File
#*********************************************************************************
PathToExist = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\July-6-2020\Models---July-2020\VISSIM\Existing'
ExistingPMfi = '20834_Existing_PM_Vehicle Travel Time Results.att'
ExistingPMfi = os.path.join(PathToExist,ExistingPMfi)
ExistingAMfi ='20834_Existing_AM_Vehicle Travel Time Results.att'
ExistingAMfi = os.path.join(PathToExist,ExistingAMfi)

def PreProcessVissimVolumes(file = ExistingAMfi):
    # Define PM file. Would later use to figure out if a file for PM 
    # This is hard coding. Would break if the name of PM file is changed.
    RefFile = ExistingPMfi
    # Read VISSIM results
    ExistingAMDat=pd.read_csv(file,sep =';',skiprows=17)
    ExistingAMDat.columns
    ExistingAMDat.rename(columns={'TRAVTM(ALL)':'VissimTT','VEHS(ALL)':'Veh'},inplace=True)
    mask=ExistingAMDat["$VEHICLETRAVELTIMEMEASUREMENTEVALUATION:SIMRUN"]=="AVG"
    ExistingAMDat = ExistingAMDat[mask]
    ExistingAMDat["TTSegNm"]=ExistingAMDat['VEHICLETRAVELTIMEMEASUREMENT'].apply(TTSegName)
    WB_TTSegs = ['WB LPGA (Clyde-Morris Blvd to Willamson Blvd)',
            'WB LPGA (Willamson Blvd to Technology Blvd)',
           'WB LPGA (Technology Blvd to I-95 NB Ramp)',
           'WB LPGA (I-95 NB Ramp to I-95 SB Ramp)',
            'WB LPGA (I-95 SB Ramp to Tomoka Rd)']
    #'4500-5400','5400-6300','6300-7200','7200-8100' are peak periods
    if (file ==ExistingPMfi):
        mask1 = (ExistingAMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100'])) & (~ExistingAMDat.TTSegNm.isin(WB_TTSegs))
        #Can include '8100-9000' as WB TT Run includes 15 min after the peak
        mask2 = (ExistingAMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100','8100-9000'])) & (ExistingAMDat.TTSegNm.isin(WB_TTSegs))
        mask = mask1 | mask2
    else:
        mask = ExistingAMDat.TIMEINT.isin(['4500-5400','5400-6300','6300-7200','7200-8100']) 
    ExistingAMDat = ExistingAMDat[mask]
    i95_NB_SB_seg = ['SB I-95 (SR40 to SB OffRamp)','SB I-95 (SB OffRamp to SB LoopRamp)',
                     'SB I-95 (SB LoopRamp to SB On-Ramp)','SB I-95 (SB On-Ramp to US92)','NB I-95 (US92 to NB OffRamp)',
                     'NB I-95 (NB OffRamp to NB LoopRamp)','NB I-95 ( NB LoopRamp to NB On-Ramp)','NB I-95 (NB On-Ramp to SR40)']
    ExistingAMDat = ExistingAMDat.query('TTSegNm.isin(@i95_NB_SB_seg)')
    ExistingAMDat.TTSegNm =pd.CategoricalIndex(data=ExistingAMDat.TTSegNm,categories=i95_NB_SB_seg,ordered=True)
    ExistingAMDat.rename(columns={'Veh':'VissimVol','TIMEINT':'HourInt'},inplace=True)
    ExistingAMDat = ExistingAMDat[['TTSegNm','HourInt','VissimVol']]
    ExistingAMDat.VissimVol = ExistingAMDat.VissimVol
    return(ExistingAMDat)

ExistingAMDat = PreProcessVissimVolumes(ExistingAMfi)
ExistingPMDat = PreProcessVissimVolumes(ExistingPMfi)

ExistingAMDat= ExistingAMDat.merge(ObsDataAm,on=['TTSegNm','HourInt'],how='left')
ExistingAMDat.loc[:, 'GEH'] = np.sqrt(((ExistingAMDat.VissimVol - ExistingAMDat.ObsVol) ** 2) / (
            (ExistingAMDat.VissimVol + ExistingAMDat.ObsVol) / 2))
ExistingAMDat.loc[:,'Intersection'] = ExistingAMDat.TTSegNm.str.strip()
ExistingAMDat.loc[:,'Movement'] = ExistingAMDat.TTSegNm.str.split(expand=True)[0].str.strip()+"T"

ExistingPMDat= ExistingPMDat.merge(ObsDataPm,on=['TTSegNm','HourInt'],how='left')
ExistingPMDat.loc[:, 'GEH'] = np.sqrt(((ExistingPMDat.VissimVol - ExistingPMDat.ObsVol) ** 2) / (
            (ExistingPMDat.VissimVol + ExistingPMDat.ObsVol) / 2))
ExistingPMDat.loc[:,'Intersection'] = ExistingPMDat.TTSegNm.str.strip()
ExistingPMDat.loc[:,'Movement'] = ExistingPMDat.TTSegNm.str.split(expand=True)[0].str.strip()+"T"

ExistingAMDat = ExistingAMDat[['Intersection','Movement','HourInt','ObsVol','VissimVol','GEH']]
ExistingPMDat = ExistingPMDat[['Intersection','Movement','HourInt','ObsVol','VissimVol','GEH']]
#*********************************************************************************
# Write to excel
#*********************************************************************************

path_to_i95_tmc = r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\July-6-2020\Models---July-2020\VISSIM'
OutFi = "i-95-processed-vissim-observed-volumes.xlsx"
OutFi = os.path.join(path_to_i95_tmc,OutFi)

writer=pd.ExcelWriter(OutFi)
# Write tables on same sheet wih 2 row spacing
ExistingAMDat.to_excel(writer,'am_i95_volume')
ExistingPMDat.to_excel(writer,'pm_i95_volume')
writer.save()


