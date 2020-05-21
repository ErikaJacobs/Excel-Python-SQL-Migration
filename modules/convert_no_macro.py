# Conversion File
# Removes Macros from Excel file and saves new .xlsm version

#%%

def save_xlsm(srcfile):
    
    import win32com.client as win32
    
    try:
        if srcfile[-7:-5] = 'A' and srcfile[-9:-7] = 'B':
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

#%%
        
def convert_all():
    
    import glob
    import datetime

    # List of all files
    xlsmfiles = glob.glob("//drive/folder/folder/folder/*")
    xlsmfiles = [i for i in xlsmfiles if (i[-4:] == 'xlsm' or i[-3:] == 'xls')]
    list=[]
    
    # Clear Win32 Temp Files - Caused Some Errors
    # If you cannot run as administrator, you can delete folder manually
    
    # import os
    # import win32com
    # print(win32com.__gen_path__)
    # path = win32com.__gen_path__
    # os.remove(path)
    
    # Error messages as comments if needed for debug
    
    def xl_read():    
        for f in xlsmfiles:
            
            # Skip process for xlsm files already converted
            if '.xlsm' in f and f.replace('.xlsm', '_new.xlsm') in xlsmfiles:
                if xlsmfiles[-1]==f:
                    if len(list) == 0:
                        print('No errors in conversion!')
                        continue
                    else:
                        print('The following file(s) experienced errors:')
                        print(list)
                        continue
                else:
                    #print("{} has already been converted.".format(f[-27:-5]))
                    continue
           
            # Skip process for xls files already converted
            if '.xls' in f and f.replace('.xls', '_new.xlsm') in xlsmfiles:
                if xlsmfiles[-1]==f:
                    if len(list) == 0:
                        print('No errors in conversion!')
                        continue
                    else:
                        print('The following file(s) experienced errors:')
                        print(list)
                        continue
                else:
                    #print("{} has already been converted.".format(f[-27:-5]))
                    continue
                
            # Skip process for converted files
            elif f[-9:-5] == "_new":
                    if xlsmfiles[-1]==f:
                        if len(list) == 0:
                            print('No errors in conversion!')
                            continue
                        else:
                            print('The following file(s) experienced errors:')
                            print(list)
                            continue
                    else:
                        #print("{} is a converted file and was skipped.".format(f[-31:-5]))      
                        continue
                    
            # Convert file
            else:  
                #print('Conditions Met for {}'.format(f[-27:-5]))
                try: 
                    save_xlsm(f)
                    #print('File Converted: {}'.format(f[-27:-5]))
                except: 
                    #print("Error on {}".format(f[-27:-5]))
                    list.append(f[-27:-5])
            if xlsmfiles[-1]==f:         
                if len(list) == 0:
                    print('No errors in conversion!')
                    continue
                else:
                    print('The following file(s) experienced errors:')
                    print(list)
                    continue
            else:
                continue
    
    print('CONVERSION PROCESS STARTED AT {}'.format(datetime.datetime.now()))
    xl_read()
    print('CONVERSION PROCESS COMPLETE AT {}'.format(datetime.datetime.now()))

#%%