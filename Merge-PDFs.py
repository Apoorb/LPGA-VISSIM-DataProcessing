# -*- coding: utf-8 -*-
"""
Created on Wed Sep 18 11:13:16 2019

@author: abibeka
Purpose: Merge pdfs
"""


from PyPDF2 import PdfFileMerger
import os
import glob


os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\ToFDOT\HCS - V1\NoBuildReports')
pdf_list = glob.glob('[0-9]*.pdf')
Pdf_Dict =  {}
for p in pdf_list:
    i = int(p.split('.')[0])
    Pdf_Dict[i] = p

merger = PdfFileMerger()
for i in range(1,len(pdf_list)+1):
    merger.append(Pdf_Dict[i])

merger.write("NoBuild-HCM.pdf")
merger.close()

os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\ToFDOT\HCS - V1\BuildReport')
pdf_list = glob.glob('[0-9]*.pdf')
Pdf_Dict =  {}
for p in pdf_list:
    i = int(p.split('.')[0])
    Pdf_Dict[i] = p

merger = PdfFileMerger()
for i in range(1,len(pdf_list)+1):
    merger.append(Pdf_Dict[i])

merger.write("Build-HCM.pdf")
merger.close()




os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\ToFDOT\HCS - V1\ExistingReport')
pdf_list = glob.glob('[0-9]*.pdf')
Pdf_Dict =  {}
for p in pdf_list:
    i = int(p.split('.')[0])
    Pdf_Dict[i] = p

merger = PdfFileMerger()
for i in range(1,len(pdf_list)+1):
    merger.append(Pdf_Dict[i])

merger.write("Existing-HCM.pdf")
merger.close()




merger = PdfFileMerger()
l = os.walk(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\ToFDOT\Synchro-Results\Build')
Folders = [x[0] for x in l]
Folders = Folders[1:]
folder = Folders[0]
for folder in Folders:
    os.chdir(folder)
    pdf_list = glob.glob('[0-9]*.pdf')
    Pdf_Dict =  {}
    for p in pdf_list:
        i = int(p.split('.')[0])
        Pdf_Dict[i] = p
    for i in range(1,len(pdf_list)+1):
        merger.append(Pdf_Dict[i])

os.chdir(r'C:\Users\abibeka\OneDrive - Kittelson & Associates, Inc\Documents\LPGA\ToFDOT\Synchro-Results\Build')
merger.write("Build-Synchro.pdf")


