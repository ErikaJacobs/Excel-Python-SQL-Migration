# List Determines how many files to push through the process
# If there are files that need conversion, only those files will be replaced in SQL Server

def convert_check():
    import glob
    
    # List of all files
    allfiles = glob.glob("//drive/folder/folder/folder/*")
    allfiles = [i for i in allfiles if '~$' not in i]   
    allfiles = [i for i in allfiles if (i[-4:] == 'xlsm' or i[-3:] == 'xls')]
    
    
    newfiles = [i for i in allfiles if i[-9:] == '_new.xlsm']
    oldfiles = [i for i in allfiles if i[-9:] != '_new.xlsm']
    
    # Check For Files Needing Conversion
    for file in newfiles:
        index = newfiles.index(file)
        newfiles[index]= newfiles[index].replace('_new.xlsm','')
        newfiles[index]= newfiles[index].replace('//drive/folder/folder/folder\\filename ','')
    
    for file in oldfiles:
        index = oldfiles.index(file)
        oldfiles[index]= oldfiles[index].replace('.xlsm','')
        oldfiles[index]= oldfiles[index].replace('.xls','')
        oldfiles[index]= oldfiles[index].replace('//drive/folder/folder/folder\\filename ','')
    
    convert = []
    
    for file in oldfiles:
        if file not in newfiles:
            convert.append(file)
    return convert
        
#%%