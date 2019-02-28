# https://github.com/sbha/readdir_xl.py

import pandas as pd
import os
import glob 
import re


dir_path = '/Users/stuartharty/Documents/Data/test_dir/'
file_pattern = "test[0-9].xlsx"
file_pattern = 'sample_*.xlsx'

def dir_xl_reader(d, f):
    df_out = pd.DataFrame()
    for file in glob.glob(os.path.join(d, f)):
        df = pd.read_excel(file)
        df['file'] = file
        df['file'] = df['file'].str.replace(d, '')
        df_out = df_out.append(df, ignore_index=True)
    return(df_out)
    
df_xl = dir_xl_reader(dir_path, file_pattern)
df_xl



def dir_xl_reader(d, f):
    df = pd.DataFrame()
    for file_name in glob.glob(os.path.join(d, f)):
        sheets = pd.ExcelFile(file_name).sheet_names

        for sheet in sheets:
            df_sheet = pd.read_excel(file_name, sheet)
            df_sheet['sheet_name'] = sheet
            df_sheet['file_name'] = file_name
            df_sheet['file_name'] = df_sheet['file_name'].str.replace(d, '')
            df = df.append(df_sheet) 

    df = df[['file_name', 'sheet_name', 'cat'] + df.columns[:-3].tolist()]
    df = df.rename(columns=lambda x: re.sub('\s+','_',x)) 
    df.columns = df.columns.str.lower()       
    return(df)
    
df_xl = dir_xl_reader(dir_path, file_pattern)
df_xl.head()



def dir_xl_reader(d, f):
    df = pd.DataFrame()
    for file_name in glob.glob(os.path.join(d, f)):
        sheets = pd.ExcelFile(file_name).sheet_names

        for sheet in sheets:
            df_sheet = pd.read_excel(file_name, sheet)
            df_sheet['sheet_name'] = sheet
            df_sheet['file_name'] = file_name
            df_sheet['file_name'] = df_sheet['file_name'].str.replace(d, '')
            df = df.append(df_sheet) 

    df = df[['file_name', 'sheet_name', 'cat'] + df.columns[:-3].tolist()]
    df = df.rename(columns=lambda x: re.sub('\s+','_',x)) 
    df.columns = df.columns.str.lower() 
    df = df[df['cat'] != 'b']  
    df = df.drop('col_3', 1)
    df['col_1_plus_col_2'] = df['col_1'] + df['col_2']
    return(df)
    
df_xl = dir_xl_reader(dir_path, file_pattern)
df_xl.head()


df_xl.groupby(['file_name', 'sheet_name']).size()