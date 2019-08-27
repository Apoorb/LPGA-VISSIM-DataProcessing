# -*- coding: utf-8 -*-
"""
Created on Mon Aug 26 16:34:41 2019

@author: abibeka
"""

import win32com.client as com
import glob

Vissim = com.Dispatch("Vissim.Vissim-64.900") # Vissim 9 - 64 bit

ExistingFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Existing\*.inpx')
NoBuildFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\No Build\*.inpx')
DCDCFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCDI\*.inpx')
DCMIFile=glob.glob(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\VISSIM-Files\VISSIM - V2\Build\DCMI\*.inpx')

ListFiles = NoBuildFile+DCDCFile+DCMIFile

file = ListFiles[0]
for file in ListFiles:
    layoutFile=file.replace(".inpx",".layx")
    flag_read_additionally  = False # you can read network(elements) additionally, in this case set "flag_read_additionally" to true
    Vissim.LoadNet(file, flag_read_additionally)
    Vissim.LoadLayout(layoutFile)
    ## ========================================================================
    # Simulation
    #==========================================================================
    # 10 runs 
    Vissim.Simulation.SetAttValue('UseMaxSimSpeed', True)
    Vissim.Graphics.CurrentNetworkWindow.SetAttValue('QuickMode', True)
    #Vissim.Simulation.SetAttValue('RandSeed', cnt_Sim + 1) # Note: RandSeed 0 is not allowed
    Vissim.Simulation.RunContinuous()

## ========================================================================
# End Vissim
#==========================================================================
Vissim = None