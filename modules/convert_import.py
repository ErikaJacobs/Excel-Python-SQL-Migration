import glob
import datetime
import win32com.client as win32
import numpy as np
import pandas as pd

class convert_import:
    
    def __init__(self):
        
        # List of all files
        self.allfiles = glob.glob("//drive/folder/folder/folder/*")
        self.allfiles = [i for i in self.allfiles if (i[-4:] == 'xlsm' or i[-3:] == 'xls') and '~$' not in i]
        
        # Files needing conversion
        self.conv = []
        
        # Files being imported
        self.imp = []
        
# Check for files needing conversion
        
    def check(self):
        
        # List of all files
        allfiles = self.allfiles
        
        # List of files to convert
        conv = self.conv
        
        newfiles = []
        oldfiles = []
        
        # Dividing old and new files
        for i in allfiles:
            if i[-9:] == '_new.xlsm':
                newfiles.append(i)
            else:
                oldfiles.append(i)
        
        # Check For Files Needing Conversion
        for i in range(len(oldfiles)):
            if oldfiles[i][-4:] == 'xlsm':
                file = oldfiles[i].replace('.xlsm','_new.xlsm')
            else:
                file = oldfiles[i].replace('.xls','_new.xlsm')
            
            if file not in newfiles:
                conv.append(oldfiles[i])
                
# Converts Excel files that need conversion    
                
    def convert(self):
        
        # Deletes Macros - Avoid Corrupted Macro
        def save_xlsm(srcfile):
        
            try:
                if srcfile[-7:-5] >= '?' and srcfile[-9:-7] >= '?':
                    password = 'DESC1'
                else:
                    password = 'DESC2'  
                    
                if srcfile[-4:] == 'xlsm':
                    newfile = srcfile.replace('.xlsm', '_new.xlsm').replace("/", "\\")
                else: 
                    newfile = srcfile.replace('.xls', '_new.xlsm').replace("/", "\\")
        
                # Excel Settings        
                xlApp = win32.gencache.EnsureDispatch('Excel.Application')
                wb = xlApp.Workbooks.Open(srcfile, False, True, None, Password=password) 
                wb.Visible = False
                wb.EnableEvents = False
                wb.DisplayAlerts = False
                
                # Delete Macros
                for i in wb.VBProject.VBComponents:        
                    xlmodule = wb.VBProject.VBComponents(i.Name)
                    if xlmodule.Type in [1, 2, 3]:            
                        wb.VBProject.VBComponents.Remove(xlmodule) 
                        
                # Save new files
                xlApp.Worksheets("DATA").Activate() 
                wb.SaveAs(newfile, 52)
                
            except Exception as e:
                print(e)  
                
            finally:
                wb.Close(True); wb = None
                xlApp.Quit; xlApp = None
            
            return newfile
    
        # List of files needing conversion
        xlsmfiles = self.conv
        errors=[]
        
        # Establish end of conversion list
        if len(xlsmfiles) > 1:
            end = xlsmfiles[-1]
        else:
            end = xlsmfiles[0]
        
        # Conversion Loop
        def xl_read():    
            for f in xlsmfiles:
                try: 
                    newfile = save_xlsm(f)
                    self.imp.append(newfile)
                except:
                    ind = f.find('WORD')
                    errors.append(f[ind:])
            
            # Print out conversion errors at end        
            if f == end:         
                if len(errors) == 0:
                    print('No errors in conversion!')
                    pass
                else:
                    print('The following file(s) experienced errors:')
                    for file in errors:
                        print(file)
                    pass
            else:
                pass
        
        print(f'CONVERSION PROCESS STARTED AT {datetime.datetime.now()}')
        xl_read()
        print(f'CONVERSION PROCESS COMPLETE AT {datetime.datetime.now()}')
        
# Imports Excel Files to Python

    def import_all(self):
        
        # List of all files being imported
        if len(self.imp) > 0:
            files = self.imp

        else:
            files = [i for i in self.allfiles if '_new.xlsm' in i]

        # Setting Data Types
        self.dtype={'variable1': str, 'variable2': str, 'variable3': str, \
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
        
        dtype = self.dtype
        
        print(f'IMPORT PROCESS STARTED AT {datetime.datetime.now()}')
        
        # Bringing in first file - Create DF
        firstfile = files[0]
        self.df = pd.read_excel(firstfile, sheet_name='DATA', header=1, dtype=dtype)
        df = self.df
        
        if 'WORD' in firstfile:
            ind = firstfile.find('WORD ')+len('WORD ')
            df['variable9'].fillna('??', inplace=True)
            df['variable9']= df['variable9'].str.replace('?', '', case = False)
            df['variable9']= df['variable9'].str.replace(' ', '', case = False)
        else:
            ind = firstfile.find('WORD ')+len('WORD ')
            df['variable9'].fillna('??', inplace=True)
            df['variable9']= df['variable9'].str.replace('?', '', case = False)
            df['variable9']= df['variable9'].str.replace(' ', '', case = False)
            
        df['variable152']=firstfile[ind:ind+2]
        df['variable153']='{}'.format(firstfile[ind+2:ind+4])
        df['variable154']="{}/{}".format(firstfile[ind:ind+2],firstfile[ind+2:ind+4])
        
        # Loop to append remaining files to DF (if any)
        if len(files) > 1:
            
            for file in range(1,len(files)):
                f = files[file]
                temp = pd.read_excel(f, sheet_name='DATA', header=1, dtype=dtype)

                if 'WORD' in f:
                    ind = f.find('WORD ')+len('WORD ')
                    temp['variable9'].fillna('??', inplace=True)
                    temp['variable9']= temp['variable9'].str.replace('?', '', case = False)
                    temp['variable9']= temp['variable9'].str.replace(' ', '', case = False)
                else:
                    ind = f.find('WORD ')+len('WORD ')
                    temp['variable9'].fillna('??', inplace=True)
                    temp['variable9']= temp['variable9'].str.replace('?', '', case = False)
                    temp['variable9']= temp['variable9'].str.replace(' ', '', case = False)
                
                temp['variable152']=f[ind:ind+2]
                temp['variable153']='{}'.format(f[ind+2:ind+4])
                temp['variable154']="{}/{}".format(f[ind:ind+2],f[ind+2:ind+4])
                
                self.df = self.df.append(temp, sort = False)
                df = self.df
                
        print(f'IMPORT PROCESS COMPLETED AT {datetime.datetime.now()}')
        
# Basic Cleaning of Imported Data
    
    def clean(self):
        
        print(f'CLEANING PROCESS STARTED AT {datetime.datetime.now()}')
        
        # Reset index
        df = self.df
        df.reset_index()
    
        # Adjust variable
        if 'variable' not in df.columns:
            df['variable'] = np.nan
        
        if 'variable8' not in df.columns:
            df['variable8'] = np.nan
            
        df['variable'].fillna(df['variable8'], inplace=True)
        df['variable8']=df['variable']
        
        # Add column placeholders (In case columns are missing)
        cols = df.columns
        
        for c in self.dtype.keys():
        
            if c not in cols:
                df[c] = np.nan
                df.astype({c: self.dtype[c]})
        
        # Drop Columns - Not Needed
        for c in cols:
            if c not in self.dtype.keys():
                df=df.drop(c, axis=1)
        
        # Adjust variable7 - different format depending on year of file
        df['variable7']=df['variable7'].apply(lambda x: str(x).replace("\\","/"))
        
        try:
            df = df[df.variable7 != '////']
        except:
            pass
        
        # Remove NaN Values
        df = df.fillna(value=0) # Fill NaN values with 0
        
        # Order Columns Properly
        
        df = df[list(self.dtype.keys())]
        
        self.df = df
        
        print('CLEANING PROCESS COMPLETED AT {}'.format(datetime.datetime.now()))

#%%