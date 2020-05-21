# Basic Cleaning and Transformation of SQL Data

def Transform():
    import pyodbc
    import datetime
    
    #Creates Connection
    conn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};'
                          'Server=Server Name'
                          'Database=Database Name;'
                          'Trusted_Connection=yes;')
    cursor = conn.cursor()
    
    # Import .sql File
    queries = open('//drive/folder/folder/folder/folder/Final/SQL/SQL_Transform.sql', 'r')
    sqlFile = queries.read()
    queries.close()

    # all SQL commands (split on ';')
    sqlCommands = sqlFile.split(';')

    # Execute every command from the input file
    print('SQL TRANSFORMATIONS STARTED AT {}'.format(datetime.datetime.now()))
    for command in sqlCommands:
        try:
            cursor.execute(command)
            conn.commit()
        except:
            print("COMMAND FAILURE: '"'{}'"'".format(command))
            
    print('SQL TRANSFORMATIONS COMPLETE AT {}'.format(datetime.datetime.now()))

#%%