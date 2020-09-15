# MAIN APPLICATION FILE
# Excel ---> Python ---> SQL

def run():
    
    # Import Packages
    import modules.convert_import as m1
    import modules.sql as m2
    import modules.warning as warning

    # Convert, Import, and Clean
    c1 = m1.convert_import()
    c1.check()
    
    if len(c1.conv) < 1:
        warning.window()
    else:
        c1.convert()
    
    c1.import_all()
    c1.clean()
    
    # Output to SQL
    c2 = m2.sql(c1.df, c1.imp)
    c2.export()

#%%