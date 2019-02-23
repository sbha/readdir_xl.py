file_type = "test[0-9].xlsx"

def dir_xl_reader(d, f):
    df_out = pd.DataFrame()
    for file in glob.glob(os.path.join(d, f)):
        df = pd.read_excel(file)
        df['file'] = file
        df['file'] = df['file'].str.replace(d, '')
        df_out = df_out.append(df, ignore_index=True)
    return(df_out)
    
df_xl = dir_xl_reader(dir_path, file_type)
df_xl

