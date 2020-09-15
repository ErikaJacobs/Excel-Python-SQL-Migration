# Warning Message - Warns User of Long Processing Time for ALL files

def window():
    import win32ui
    import win32con
    import sys
    
    Title = "WARNING!"
    
    Message = "You are about to export ALL files to SQL Server. "\
    +"This process uses a large amount of processing time. "\
    +"Interruption of this process could cause future errors. "\
    +"Would you like to proceed?"
    
    # A prompt for Yes/No/Cancel
    response = win32ui.MessageBox(Message, Title, win32con.MB_YESNO)
    if response == win32con.IDNO:
        win32ui.MessageBox("Process Cancelled")
        sys.exit()
    
#%%