## Introduction

Python can quickly and easily aggregrate data from multiple sheets from multiple Excel workbooks into a single data frame. A custom function combining functions from the `Pandas` package provides a template that can be adapted and expanded to handle many different data specific needs. 

This exercise is also done using [R's Tidyverse](https://github.com/sbha/readdir_xl).

## Usage

Import the packages that have the functions that do all the heavy lifting:

```
import pandas as pd
import os
import glob 
import re
```

Define the target directory as a variable. This is the path to the directory where the Excel files are stored:

`dir_path <- "~/path/to/test_dir/"`  

In this example all the Excel files we're interested in have the same naming convention, so we can use a matching pattern to ensure we're reading only the files we need. This is useful if there are files in the same directory that we don't want to read. In this case the files are named `sample_2019-01-09.xlsx`, `sample_2019-01-15.xlsx` and so on, but we could simply use `xlsx` if we knew we needed every Excel file, or leave that arguement empty if we knew the directory contained only Excel files. The pattern can be defined as a variable:

`file_pattern <- "^sample_*.xlsx" `   

This pattern will match file names that begin with `sample_` and end with the file extension `.xlsx`. You can check that this matches all the files you're expecting with:

`glob.glob(os.path.join(dir_path, f))`

Next we can write the custom function that will read the individual Excel files and combine the different sheets within each workbook into a single data frame. The two function inputs will be the path to the files, which was defined earlier with `dir_path` and then the file name pattern `file_pattern`. The function will loop through the files in the directory that match the file name pattern, get the individual sheet neames from each file with `pd.ExcelFile(file_name).sheet_names`, then import the data from each sheet with `pd.read_excel(file_name, sheet)`. The file names and sheet names are added as columns, and along with the column `cat` are moved to the leftmost columns to improve readability of the data frame:

```
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
    return(df)
```    

With the custom function `dir_xl_reader()` defined, we can now use it to aggregate all the sheets from all the individual files matching `file_pattern` in the directory into a single data frame:

```
df_xl = dir_xl_reader(dir_path, file_pattern)
>>> df_xl.head()
                file_name sheet_name cat  Col 1  Col 2  Col 3  Col 4
0  sample_2019-02-02.xlsx     Sheet1   b    5.0    8.0    1.0    NaN
1  sample_2019-02-02.xlsx     Sheet1   b    8.0    9.0    2.0    NaN
2  sample_2019-02-02.xlsx     Sheet1   c    9.0    1.0    5.0    NaN
3  sample_2019-02-02.xlsx     Sheet1   a    1.0    5.0    3.0    NaN
4  sample_2019-02-02.xlsx     Sheet1   b    2.0    8.0    NaN    NaN
```

To get a better sense of everything in the data frame, we can use `df_xl.groupby(['file_name', 'sheet_name']).size()` to see the number of observations by file and sheet name:

```
>>> df_xl.groupby(['file_name', 'sheet_name']).size()
file_name               sheet_name
sample_2019-01-09.xlsx  Sheet1        12
                        Sheet2        11
sample_2019-01-15.xlsx  Sheet1        12
                        Sheet2         8
                        Sheet3         7
sample_2019-02-02.xlsx  Sheet1        17
                        Sheet2         8
                        Sheet3         5
dtype: int64
```

In this example, not every file has the same sheets or columns. The files need not have exactly the same structure; only `sample_2019-01-15.xlsx` has `Col 4`, and `sample_2019-01-09.xlsx` has only two sheets. 

This custom function aggregates every file, but we can improve the output by extending the function to handle more specific needs of this data. For our example, the column names can be reformatted so that they are always in a more `Python` friendly format. In the modified function below, all column names are converted to lower case and spaces are replaced with a single underscores:

```
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
```

Going further, if the original Excel files contains data you don't need, you could remove it by filtering or by dropping specific columns. Or a column could be created using with something like the following:

```
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
```

In this expanding function, we're removing any rows where column `cat` equals `b`, dropping column `col_3`, and then creating a new column that adds columns `col_1` and `col_2`. Running the same Excel files through the updated custom function gives us a different data frame, more specific to our needs:

```
df_xl = dir_xl_reader(dir_path, file_pattern)
>>> df_xl.head()
                 file_name sheet_name cat  col_1  col_2  col_4  col_1_plus_col_2
2   sample_2019-02-02.xlsx     Sheet1   c    9.0    1.0    NaN              10.0
3   sample_2019-02-02.xlsx     Sheet1   a    1.0    5.0    NaN               6.0
6   sample_2019-02-02.xlsx     Sheet1   c    9.0    8.0    NaN              17.0
7   sample_2019-02-02.xlsx     Sheet1   a    1.0    9.0    NaN              10.0
10  sample_2019-02-02.xlsx     Sheet1   c    4.0    5.0    NaN               9.0
```

The modifications with this example might seem trivial and something that can be done seperately after the data has been aggregrated, which they can, of course, but they quickly become useful when dealing with a large number of files that might get near a machine's memory limits or as part of a process that will be repeated. 

Finally, if the files aren't `.xlsx`, a similar method can be used for `.csv` or other delimited files using functions from the Pandas library:

```
def dir_reader(d, f):
    df_out = pd.DataFrame()
    for file in glob.glob(os.path.join(d, f)):
        df = pd.read_csv(file)
        df['file'] = file
        df['file'] = df['file'].str.replace(d, '')
        df_out = df_out.append(df, ignore_index=True)
    return(df_out)

file_type = "test[0-9].csv"    
df = dir_reader(dir_path, file_type)

```


### Summary

Using functions from the `Pandas` package we can define a custom function that can combine multiple Excel files into a single data frame. We can modifiy and extend that custom function further to handle more specific needs for formatting, reorganizing, and combining the data. This is a process that can be adapted as needed. 
