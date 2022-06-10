import pandas as pd
import openpyxl
import pprint
import numpy as np
import shutil
import os
import argparse

parser = argparse.ArgumentParser(description = "This script runs automatically for the number of students studying weekly")

parser.add_argument("--YSOF", "-YS", required = True, metavar ="",
                    help = "Input the original list of students")

parser.add_argument("--indir", "-i", required = True, metavar ="",
                    help = "absolute path for the xlsx file containing information about  the number of students studying")

parser.add_argument("--Zoom" , "-z", required = True, metavar = "",
                    help = "Input the name of zoom. Ex: Zoom23")

parser.add_argument("--LuongGia" , "-lg", required = True, metavar = "",
                    help = "Input the name of LG. Ex: LG23")

parser.add_argument("--information" , "-infor", metavar = "", type=int,
                    help = "Overview information about dataframe")

parser.add_argument("--clean" , "-cl", metavar = "", type=int,
                    help = "Clean data")
args   = parser.parse_args()
#__________________________________________________________INPUT WORKING DIRECTORY__________________________________________________________
prefix         = "YSOF"
prefix_diemdanh="DSHV_YSOF_diemdanh"
#__________________________________________________________IMPORT DATAFRAME FROM EXCEL__________________________________________________________
dshv           = pd.read_excel(args.YSOF)
dshv_excel     = pd.ExcelFile (args.YSOF)

filexcel       = pd.ExcelFile (args.indir)
sheet_name     = filexcel.sheet_names # List of sheet names in file excel
zoom_name      = [sheet_name for sheet_name in sheet_name if sheet_name.startswith("Điểm danh")]
lg_name        = [sheet_name for sheet_name in sheet_name if sheet_name.startswith("Lượng giá")]

df             = [pd.read_excel(filexcel, sheet_name= sheet) for sheet in sheet_name]

all_df         = []
all_df.append(dshv)
all_df         = all_df + df 
all_name       = dshv_excel.sheet_names + sheet_name
#__________________________________________________________OVERVIEW DATAFRAME__________________________________________________________
# Overview of dataframe
if args.information == 1:
    for i in range(len(all_df)):
        print (f"Checking {all_name[i]} ")
        all_df[i].info()
        print ("\n")

# Check duplicate
if args.information == 2:
    for i in range(len(all_df)):
        print (f"Checking {all_name[i]}")
        if not all_df[i]["MSHV"].squeeze().is_unique:
            print (f"{all_name[i]} contains duplicate values" + "\n")
            '''Out put the duplicate rows'''          
            MSHV_duplicate_Series = all_df[i].duplicated(subset=["MSHV"], keep = False)
            MSHV_duplicate        = all_df[i][MSHV_duplicate_Series]
            print (MSHV_duplicate)
                       
# Check inconsistent values
if args.information == 3:
    categories                   = {"Categories":[1,"KP","P"]}
    categories_df                = pd.DataFrame.from_dict(categories)

    for i in range(len(zoom_name)): 
        '''Check information about Zoom'''          
        print (f"Checking {sheet_name[i]}")
        inconsistent_zoom        = set(df[i][args.Zoom]).difference(categories_df["Categories"])
        inconsistent_zoom_series = df[i][args.Zoom].isin(inconsistent_zoom)
        inconsistent_zoom_row    = df[i][inconsistent_zoom_series]
        print (inconsistent_zoom_row)

        '''Check information about LG'''
        print (f"Checking {lg_name[i]}")
        inconsistent_lg         = set(df[i+1][args.LuongGia]).difference(categories_df["Categories"])
        inconsistent_lg_series  = df[i+1][args.LuongGia].isin(inconsistent_lg)
        inconsistent_lg_row     = df[i+1][inconsistent_lg_series]
        print (inconsistent_lg_row)
 
#__________________________________________________________CLEANING DATA IN DATAFRAME__________________________________________________________
if args.clean == 2:
    for i in range(len(df)):
        df[i].drop_duplicates(inplace=True, keep="first")
        df[i][args.Zoom]     = df[i][args.Zoom].str.strip()
        '''Uppercase the information of Zoom'''
        df[i][args.Zoom]     = df[i][args.Zoom].str.upper()
#__________________________________________________________MERGE DATAGRAME_____________________________________________________________________
identification               = pd.merge(dshv,df[0],on='MSHV',how='left') 
identification.drop(identification.columns[identification.columns.str.contains('unnamed: 2',case = False)],axis = 1, inplace = True) # remove "unnamed: 2" cloumn at the end"
identification[args.Zoom]    = identification[args.Zoom].fillna(".") 
identification               = pd.merge(identification,df[1],on='MSHV',how='left') 

identification[args.LuongGia]= np.where(identification[args.Zoom]=='KP','KP',
                                        np.where(identification[args.Zoom]=='P','P',
                                        np.where(identification[args.Zoom]=='.','.',
                                        np.where((identification[args.Zoom]==1) & (identification[args.LuongGia].isna()),'KP',1)))) 

zoom = identification.groupby(args.Zoom).size()
print ("The summary of students studying: ", zoom, "\n")

LG   = identification.groupby(args.LuongGia).size()
print ("The summary of students filling form: ", LG, "\n")

identification["result"]    = np.where(identification[args.LuongGia]=='1','complete',
                                        np.where((identification[args.Zoom]=='KP') | (identification[args.LuongGia]=='KP'),'incompletement',
                                        np.where(identification[args.Zoom]=='P','P','.')))
identification              = identification.replace('.',np.nan) 

#__________________________________________________________CLEANING DATA IN DATAFRAME__________________________________________________________
directory = os.path.dirname(args.YSOF)
if not os.path.exists(f"{directory}/2.result"):
    os.makedirs(f"{directory}/2.result")
# ---------------------------------------------
if not os.path.exists(f"{directory}/{prefix_diemdanh}_{args.Zoom}.xlsx"):
    result = identification.to_excel(f"{directory}/2.result/{prefix_diemdanh}_{args.Zoom}.xlsx",index=False)
print (result)