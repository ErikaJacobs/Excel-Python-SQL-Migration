# MAIN APPLIvariable1ION FILE
# Excel --> Python --> SQL

def run():
    
    # Import Packages
    import modules.convert_check as m1
    import modules.convert_no_macro as m2
    import modules.excel_import as m3
    import modules.clean as m4
    import modules.sql_export as m5
    import modules.sql_transform as m6
    import modules.sql_load as m7
    import modules.warning as m8

    # Procedure
    convert = m1.convert_check()
    
    # One or more updated files
    if len(convert) != 0:
        m2.convert_all()
        df = m3.import_some(convert)
        df = m4.clean_some(df)
        m5.SQL_Exp_Some(df)

    # Updating all files
    else:
        m8.warning()
        m2.convert_all()
        df = m3.import_all()
        df = m4.clean_all(df)
        m5.SQL_Exp_All(df)

    # SQL Queries
    m6.Transform()
    m7.Load()

#%%
    
#https://dev.to/codemouse92/dead-simple-python-project-structure-and-imports-38c6
#https://www.reddit.com/r/learnpython/comments/5oxkfv/python_file_naming_standard/
# https://www.internalpointers.com/post/modules-and-packages-create-python-project
