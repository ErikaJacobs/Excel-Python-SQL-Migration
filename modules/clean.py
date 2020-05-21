# Cleans Pandas P&L df imported

#%%

def clean_all(df):
    
    import pandas as pd
    import datetime
    
    print('CLEANING PROCESS STARTED AT {}'.format(datetime.datetime.now()))
    
    df.reset_index()
    df['variable'].fillna(df['variable8'], inplace=True)
    df['variable8']=df['variable']
    df=df.drop(['variable', 'variable11', 'variable12', 'variable13', 'variable14', 'variable15', 'Unnamed: 151', 'Unnamed: 152',
            'Unnamed: 153', 'Unnamed: 154', 'Unnamed: 155'], axis=1)
    df['variable7']=df['variable7'].apply(lambda x: str(x).replace("\\","/"))
    df['variable16']=df['variable7'].apply(lambda x: str(x).replace("\\","/"))
    df = df[df.variable7 != '////']
    df = df[df.variable16 != '0/0/0/0/0']
    df=df.drop(['variable16'], axis=1)
    df = df.fillna(value=0) # Fill NaN values with 0
    
    print('CLEANING PROCESS COMPLETED AT {}'.format(datetime.datetime.now()))
    
    return df

#%%
    
def clean_some(df):
    
    import pandas as pd
    import numpy as np
    import datetime
    
    print('CLEANING PROCESS STARTED AT {}'.format(datetime.datetime.now()))
    
    
    df.reset_index()

    # Adjust variable - different format depending on variable of file
    if 'variable' not in df.columns:
        df['variable'] = np.nan
    
    if 'variable8' not in df.columns:
        df['variable8'] = np.nan
        
    df['variable'].fillna(df['variable8'], inplace=True)
    df['variable8']=df['variable']
    
    # ADD COLUMN PLACEHOLDER FOR VAR FIELDS TO MATCH SQL FORMAT
    
    cols = ['variable2', 'variable3','variable4','variable5','variable6', 'variable7', 'variable149',
            'variable150', 'variable151', 'variable155', 'variable156']
    
    for c in cols:
    
        colstypes = ['str', 'str','str','str','str', 'str', 'float64',
            'float64', 'float64', 'float64', 'float64']
    
        if c not in df.columns:
            df[c] = np.nan
            df.astype({c: colstypes[cols.index(c)]})
    
    
    # LOOP TO DROP FIELDS
    fields = ['variable', 'variable11', 'variable12', 'variable13', 'variable14', 'variable15', 'Unnamed: 151', 'Unnamed: 152',
              'Unnamed: 153', 'Unnamed: 154', 'Unnamed: 155']
    
    for field in fields:
        try:
            df=df.drop(field, axis=1)
        except:
            pass
    
    # Adjust variable1 6 - different format depending on variable of file
    try:  
        df['variable7']=df['variable7'].apply(lambda x: str(x).replace("\\","/"))
    except:
        pass
    
    try:
        df['variable16']=df['variable7'].apply(lambda x: str(x).replace("\\","/"))
    except:
        pass
    
    try:
        df = df[df.variable7 != '////']
    except:
        pass
    
    try:
        df = df[df.variable16 != '0/0/0/0/0']
    except:
        pass
    
    try:    
        df=df.drop(['variable16'], axis=1)
    except:
        pass
    
    # Remove NaN Values
    df = df.fillna(value=0) # Fill NaN values with 0
    
    print('CLEANING PROCESS COMPLETED AT {}'.format(datetime.datetime.now()))
    
    return df