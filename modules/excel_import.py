# Imports Excel Files to Python

def import_all():
    
    import datetime
    import numpy as np
    import pandas as pd
    import glob
    
    # List of all files
    allfiles = glob.glob("//drive/folder/folder/folder/*")
    allfiles = [i for i in allfiles if i[-4:] == 'xlsm']
    allfiles = [i for i in allfiles if '~$' not in i]
    
    # Setting Data Types
    dtype={'variable1': str, 'variable2': str, 'variable3': str, \
                       'variable4': str, 'variable5': str, 'variable6': str, 'variable7': str, 'variable8': str, 'variable9': str, \
                       'variable10': str, 'variable11': str, 'variable12': str, 'variable13': np.float64, 'variable14': np.float64, 'variable15': np.float64, \
                       'variable16': np.float64, 'variable17': np.float64, 'variable18': np.float64, 'variable19': np.float64, 'variable20': np.float64, 'variable21': np.float64, \
                       'variable22': np.float64, 'variable23': np.float64, 'variable24': np.float64, 'variable25': np.float64, 'variable26': np.float64, \
                       'variable27': np.float64, 'variable28': np.float64, 'variable29': np.float64, 'variable30': np.float64, \
                       'variable31': np.float64, 'variable32': np.float64, 'variable33': np.float64, 'variable34': np.float64, \
                       'variable35': np.float64, 'variable36': np.float64, 'variable37': np.float64, 'variable38': np.float64, \
                       'variable39': np.float64, 'variable40': np.float64, 'variable41': np.float64, 'variable42': np.float64, \
                       'variable43': np.float64, 'variable44': np.float64, 'variable45': np.float64, \
                       'variable46': np.float64, 'variable47': np.float64, 'variable48': np.float64, \
                       'variable49': np.float64, 'variable50': np.float64, 'variable51': np.float64, 'variable52': np.float64, \
                       'variable53': np.float64, 'variable54': np.float64, 'variable55': np.float64, \
                       'variable56': np.float64, 'variable57': np.float64, 'variable58': np.float64, 'variable59': np.float64, \
                       'variable60': np.float64, 'variable61': np.float64, 'variable62': np.float64, 'variable63': np.float64, \
                       'variable64': np.float64, 'variable65': np.float64, 'variable66': np.float64, \
                       'variable67': np.float64, 'variable68': np.float64, 'variable69': np.float64, \
                       'variable70': np.float64, 'variable71': np.float64, 'variable72': np.float64, \
                       'variable73': np.float64, 'variable74': np.float64, 'variable75': np.float64, 'variable76': np.float64, \
                       'variable77': np.float64, 'variable78': np.float64, 'variable79': np.float64, 'variable80': np.float64, \
                       'variable81': np.float64, 'variable82': np.float64, 'variable83': np.float64, 'variable84': np.float64, \
                       'variable85': np.float64, 'variable86': np.float64, 'variable87': np.float64, 'variable88': np.float64, \
                       'variable89': np.float64, 'variable90': np.float64, 'variable91': np.float64, 'variable92': np.float64, 'variable93': np.float64, 'variable94': np.float64, \
                       'variable95': np.float64, 'variable96': np.float64, 'variable97': np.float64, 'variable98': np.float64, 'variable99': np.float64, \
                       'variable100': str, 'variable101': np.float64, 'variable102': np.float64, 'variable103': np.float64, 'variable104': np.float64, \
                       'variable105': np.float64, 'variable106': np.float64, 'variable107': np.float64, 'variable108': np.float64, 'variable109': np.float64, 'variable110': np.float64, \
                       'variable111': np.float64, 'variable112': np.float64, 'variable113': np.float64, 'variable114': np.float64, 'variable115': np.float64, 'variable116': np.float64, \
                       'variable117': np.float64, 'variable118': np.float64, 'variable119': np.float64, 'variable120': np.float64, 'variable121': np.float64, \
                       'variable122': np.float64, 'variable123': np.float64, 'variable124': np.float64, 'variable125': np.float64, 'variable126': np.float64, \
                       'variable127': np.float64, 'variable128': np.float64, 'variable129': np.float64, 'variable130': np.float64, 'variable131': np.float64, \
                       'variable132': np.float64, 'variable133': np.float64, 'variable134': np.float64, 'variable135': np.float64, \
                       'variable136': np.float64, 'variable137': np.float64, 'variable138': np.float64, 'variable139': np.float64, \
                       'variable140': np.float64, 'variable141': np.float64, 'variable142': np.float64, 'variable143': np.float64, \
                       'variable144': np.float64, 'variable145': np.float64, 'variable146': np.float64, 'variable147': np.float64, \
                       'variable148': np.float64, 'variable149': np.float64, 'variable150': np.float64, \
                       'variable151': np.float64, 'variable152': str, 'variable153': str, 'variable154': str, \
                       'variable155': np.float64, 'variable156': np.float64}
    
    start = 0
    print('IMPORT PROCESS STARTED AT {}'.format(datetime.datetime.now()))
    
    for f in allfiles:
    
        if f[-4:] == 'xlsm' and f[-8:] != 'new.xlsm':
            start = start + 1
            continue
        
        if f[-4:] == 'xlsm' and f[-8:] == 'new.xlsm':
            if f==allfiles[start]:
                df = pd.read_excel(f, sheet_name='DATA', header=1, dtype=dtype)
                df['variable152']=f[-13:-11]
                df['variable153']='20{}'.format(f[-11:-9])
                df['variable154']="{}/20{}".format(f[-13:-11],f[-11:-9])
                if f[-18:-14] == 'DESC':
                    df['variable9'].fillna('??', inplace=True)
                    df['variable9']= df['variable9'].str.replace('0', '??', case = False)
                    df['variable9']= df['variable9'].str.replace(' ', '??', case = False)
                else:
                    df['variable9'].fillna('??', inplace=True)
                    df['variable9']= df['variable9'].str.replace('0', '??', case = False)
                    df['variable9']= df['variable9'].str.replace(' ', '??', case = False)
                start = 0
            else:
                temp = pd.read_excel(f, sheet_name='DATA', header=1, dtype=dtype)
                temp['variable152']=f[-13:-11]
                temp['variable153']='20{}'.format(f[-11:-9])
                temp['variable154']="{}/20{}".format(f[-13:-11],f[-11:-9])
                if f[-18:-14] == 'DESC':
                    temp['variable9'].fillna('??', inplace=True)
                    temp['variable9']= temp['variable9'].str.replace('0', '??', case = False)
                    temp['variable9']= temp['variable9'].str.replace(' ', '??', case = False)
                else:
                    temp['variable9'].fillna('??', inplace=True)
                    temp['variable9']= temp['variable9'].str.replace('0', '??', case = False)
                    temp['variable9']= temp['variable9'].str.replace(' ', '??', case = False)
                df = df.append(temp, sort=False) 
            continue
    
        else:
            start = start + 1
            continue
            
    print('IMPORT PROCESS COMPLETED AT {}'.format(datetime.datetime.now()))
    return df
    
#%%

def import_some(convert):
    
    import datetime
    import numpy as np
    import pandas as pd
    
    allfiles = []

    for file in convert:
        allfiles.append("//drive/folder/folder/folder/filename {}_new.xlsm".format(file))
    
    # Setting Data Types
    dtype={'variable1': str, 'variable2': str, 'variable3': str, \
                       'variable4': str, 'variable5': str, 'variable6': str, 'variable7': str, 'variable8': str, 'variable9': str, \
                       'variable10': str, 'variable11': str, 'variable12': str, 'variable13': np.float64, 'variable14': np.float64, 'variable15': np.float64, \
                       'variable16': np.float64, 'variable17': np.float64, 'variable18': np.float64, 'variable19': np.float64, 'variable20': np.float64, 'variable21': np.float64, \
                       'variable22': np.float64, 'variable23': np.float64, 'variable24': np.float64, 'variable25': np.float64, 'variable26': np.float64, \
                       'variable27': np.float64, 'variable28': np.float64, 'variable29': np.float64, 'variable30': np.float64, \
                       'variable31': np.float64, 'variable32': np.float64, 'variable33': np.float64, 'variable34': np.float64, \
                       'variable35': np.float64, 'variable36': np.float64, 'variable37': np.float64, 'variable38': np.float64, \
                       'variable39': np.float64, 'variable40': np.float64, 'variable41': np.float64, 'variable42': np.float64, \
                       'variable43': np.float64, 'variable44': np.float64, 'variable45': np.float64, \
                       'variable46': np.float64, 'variable47': np.float64, 'variable48': np.float64, \
                       'variable49': np.float64, 'variable50': np.float64, 'variable51': np.float64, 'variable52': np.float64, \
                       'variable53': np.float64, 'variable54': np.float64, 'variable55': np.float64, \
                       'variable56': np.float64, 'variable57': np.float64, 'variable58': np.float64, 'variable59': np.float64, \
                       'variable60': np.float64, 'variable61': np.float64, 'variable62': np.float64, 'variable63': np.float64, \
                       'variable64': np.float64, 'variable65': np.float64, 'variable66': np.float64, \
                       'variable67': np.float64, 'variable68': np.float64, 'variable69': np.float64, \
                       'variable70': np.float64, 'variable71': np.float64, 'variable72': np.float64, \
                       'variable73': np.float64, 'variable74': np.float64, 'variable75': np.float64, 'variable76': np.float64, \
                       'variable77': np.float64, 'variable78': np.float64, 'variable79': np.float64, 'variable80': np.float64, \
                       'variable81': np.float64, 'variable82': np.float64, 'variable83': np.float64, 'variable84': np.float64, \
                       'variable85': np.float64, 'variable86': np.float64, 'variable87': np.float64, 'variable88': np.float64, \
                       'variable89': np.float64, 'variable90': np.float64, 'variable91': np.float64, 'variable92': np.float64, 'variable93': np.float64, 'variable94': np.float64, \
                       'variable95': np.float64, 'variable96': np.float64, 'variable97': np.float64, 'variable98': np.float64, 'variable99': np.float64, \
                       'variable100': str, 'variable101': np.float64, 'variable102': np.float64, 'variable103': np.float64, 'variable104': np.float64, \
                       'variable105': np.float64, 'variable106': np.float64, 'variable107': np.float64, 'variable108': np.float64, 'variable109': np.float64, 'variable110': np.float64, \
                       'variable111': np.float64, 'variable112': np.float64, 'variable113': np.float64, 'variable114': np.float64, 'variable115': np.float64, 'variable116': np.float64, \
                       'variable117': np.float64, 'variable118': np.float64, 'variable119': np.float64, 'variable120': np.float64, 'variable121': np.float64, \
                       'variable122': np.float64, 'variable123': np.float64, 'variable124': np.float64, 'variable125': np.float64, 'variable126': np.float64, \
                       'variable127': np.float64, 'variable128': np.float64, 'variable129': np.float64, 'variable130': np.float64, 'variable131': np.float64, \
                       'variable132': np.float64, 'variable133': np.float64, 'variable134': np.float64, 'variable135': np.float64, \
                       'variable136': np.float64, 'variable137': np.float64, 'variable138': np.float64, 'variable139': np.float64, \
                       'variable140': np.float64, 'variable141': np.float64, 'variable142': np.float64, 'variable143': np.float64, \
                       'variable144': np.float64, 'variable145': np.float64, 'variable146': np.float64, 'variable147': np.float64, \
                       'variable148': np.float64, 'variable149': np.float64, 'variable150': np.float64, \
                       'variable151': np.float64, 'variable152': str, 'variable153': str, 'variable154': str, \
                       'variable155': np.float64, 'variable156': np.float64}
    
    start = 0
    print('IMPORT PROCESS STARTED AT {}'.format(datetime.datetime.now()))
    
    for f in allfiles:
    
        if f[-4:] == 'xlsm' and f[-8:] != 'new.xlsm':
            start = start + 1
            continue
        
        if f[-4:] == 'xlsm' and f[-8:] == 'new.xlsm':
            if f==allfiles[start]:
                df = pd.read_excel(f, sheet_name='DATA', header=1, dtype=dtype)
                df['variable152']=f[-13:-11]
                df['variable153']='20{}'.format(f[-11:-9])
                df['variable154']="{}/20{}".format(f[-13:-11],f[-11:-9])
                if f[-18:-14] == 'DESC':
                    df['variable9'].fillna('??', inplace=True)
                    df['variable9']= df['variable9'].str.replace('0', '??', case = False)
                    df['variable9']= df['variable9'].str.replace(' ', '??', case = False)
                else:
                    df['variable9'].fillna('??', inplace=True)
                    df['variable9']= df['variable9'].str.replace('0', '??', case = False)
                    df['variable9']= df['variable9'].str.replace(' ', '??', case = False)
                start = 0
            else:
                temp = pd.read_excel(f, sheet_name='DATA', header=1, dtype=dtype)
                temp['variable152']=f[-13:-11]
                temp['variable153']='20{}'.format(f[-11:-9])
                temp['variable154']="{}/20{}".format(f[-13:-11],f[-11:-9])
                if f[-18:-14] == 'DESC':
                    temp['variable9'].fillna('??', inplace=True)
                    temp['variable9']= temp['variable9'].str.replace('0', '??', case = False)
                    temp['variable9']= temp['variable9'].str.replace(' ', '??', case = False)
                else:
                    temp['variable9'].fillna('??', inplace=True)
                    temp['variable9']= temp['variable9'].str.replace('0', '??', case = False)
                    temp['variable9']= temp['variable9'].str.replace(' ', '??', case = False)
                df = df.append(temp, sort=False) 
            continue
    
        else:
            start = start + 1
            continue
            
    print('IMPORT PROCESS COMPLETED AT {}'.format(datetime.datetime.now()))
    return df
    