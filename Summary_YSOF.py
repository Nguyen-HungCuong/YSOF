import pandas as pd
import openpyxl
import pprint
import numpy as np
import shutil
import os
import argparse

parser = argparse.ArgumentParser(description = "This script is used to get the overview results of students in a year")

parser.add_argument("--indir", "-i", required = True, metavar ="",
                    help = "absolute path for the xlsx file containing information about  the number of students studying weekly")

parser.add_argument("--standard_study" , "-st", metavar = "", type=int, required = True,
                    help = "The number of completement")

parser.add_argument("--absence" , "-abs", metavar = "", type=int, required = True,
                    help = "The number of absence")

args   = parser.parse_args()

prefix_ds = "DSHV_YSOF"
prefix = "YSOF"
info_stu=['STT','Saint-Full Name','Student ID','Email','Registration']
info_tonghop=['Complete', 'Absent Permission', 'Absent Non-permission','Not Filling Form']

filexcel = pd.ExcelFile(args.indir)
summary  = pd.read_excel(filexcel)

zoom_extract = [zoom for zoom in summary.columns if zoom.startswith('Zoom')]  
lg_extract   = [lg for lg in summary.columns if lg.startswith('Lượng giá')] 
kq_extract   = [kq for kq in summary.columns if kq.startswith('Kết quả')]

#_______________________________________________________________________________
summary['Complete'] = 0 
for index, row in summary.iterrows():
    for kq in kq_extract:
        if summary.loc[index,kq] == "complete":
            summary.loc[index,'Complete'] +=1

summary['Absent Permission'] = 0
for index, row in summary.iterrows():
    for zoom in zoom_extract:        
        if summary.loc[index,zoom] == 'P':
            summary.loc[index,'Absent Permission'] += 1

summary['Absent Non-permission'] = 0
for index, row in summary.iterrows():
    for zoom in zoom_extract:        
        if summary.loc[index,zoom] == 'KP':
            summary.loc[index,'Absent Non-permission'] += 1

summary['Not Filling Form'] = 0
for index, row in summary.iterrows():
    for i in range(len(zoom_extract)):
        if summary.loc[index,zoom_extract[i]] == 1:
            if summary.loc[index,lg_extract[i]] == "KP":
                summary.loc[index,'Not Filling Form'] += 1

#_______________________________________________________________________________
# SUMMARY
'''Ouput the final result'''
kq=info_stu + kq_extract + info_tonghop
summary_out=summary[kq]

summary['Total Absence'] = summary['Absent Permission'] + summary['Absent Non-permission']
demand_standard          = summary[summary['Complete'] > 14 & summary['Total Absence'] < 3]

#_______________________________________________________________________________
# Export to excel
directory = os.path.dirname(args.indir)
with pd.ExcelWriter(f"{directory}/2.summary/{prefix_ds}_summary.xlsx",engine="openpyxl") as writer:
    summary_out.to_excel(writer,sheet_name='Summary',index=False)
    demand_standard.to_excel(writer,sheet_name='Demand_Standard',index=False)
