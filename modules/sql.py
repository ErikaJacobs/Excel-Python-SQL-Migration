import pandas as pd
import datetime
from sqlalchemy import create_engine, event

class sql:
    
    def __init__(self, df, imp):
        self.driver='ODBC+Driver+17+for+SQL+Server'
        self.server='ServerName'
        self.database='DBName'
        self.conn = create_engine(f'mssql+pyodbc://{self.server}/{self.database}?driver={self.driver}?trusted_connection=yes')
        self.cursor = self.conn.connect()
        self.df = df
        self.imp = imp
        
        @event.listens_for(self.conn, "before_cursor_execute")
        def receive_before_cursor_execute(conn, cursor, statement, params, context, executemany):
            if executemany:
                cursor.fast_executemany = True
    
    def export(self):
        
        # Bring in Connection
        cursor = self.cursor
        conn = self.conn
        
        print(f'SQL EXPORT PROCESS STARTED AT {datetime.datetime.now()}')
        
        #Inserting SOME data into SQL from Python
        if len(self.imp) > 0:
            
            self.df.to_sql('Table2', conn, schema = 'dbo', if_exists = 'replace', index = False)
        
            # DELETE OLD RECORDS 
            cursor.execute('''DELETE FROM dbo.Table1 \
                           WHERE CONCAT(variable1, variable5) in(SELECT distinct CONCAT(variable1, variable5) FROM dbo.Table2)''')
            
            # CREATE UNION BETWEEN HISTORY AND NEW RECORDS
            cursor.execute('''SELECT * \
                           INTO dbo.Table3 \
                           FROM dbo.Table1 \
                           UNION ALL \
                           SELECT * \
                           FROM dbo.Table2''')
            
            # UPDATE TABLE WITH UNION
            cursor.execute('''DROP TABLE dbo.Table1 \
                           SELECT * \
                           INTO dbo.Table1 \
                           FROM dbo.Table3''')
            
            # DELETE TEMP TABLES
            cursor.execute("IF OBJECT_ID('dbo.Table3', 'U') IS NOT NULL DROP TABLE dbo.Table3;")        
            cursor.execute("IF OBJECT_ID('dbo.Table2', 'U') IS NOT NULL DROP TABLE dbo.Table2;") 
        
        #Inserting ALL data into SQL from Python
        else:
             self.df.to_sql('Table1', conn, schema='dbo', if_exists='replace', index=False)
            
        
        # UPDATE VARCHAR DATA TYPES
        
        varchars = ['variable1', 'variable2', 'variable3', 'Yariable4', 'variable5', 'variable6']
        
        for var in varchars:
            cursor.execute(f'ALTER TABLE dbo.Table1 ALTER COLUMN [{var}] VARCHAR(30);')           
        
        # CREATE/UPDATE INDEXES
        cursor.execute("CREATE INDEX variable1 ON dbo.Table1 (variable1);") 
        cursor.execute("CREATE INDEX variable2 ON dbo.Table1 (variable2);")
        cursor.execute("CREATE INDEX variable3 ON dbo.Table1 (variable3);")
        cursor.execute("CREATE INDEX variable4 ON dbo.Table1 (variable4);")
        
        # UPDATE BACK-UP TABLE
        cursor.execute("IF OBJECT_ID('dbo.Table1_Backup', 'U') IS NOT NULL DROP TABLE dbo.Table1_Backup;")
        
        cursor.execute('''SELECT * \
                       INTO dbo.Table1_Backup \
                       FROM dbo.Table1''')
        
        print(f'SQL EXPORT PROCESS COMPLETED AT {datetime.datetime.now()}')
        
        self.transform_load()

    def transform_load(self):
        
        #Creates Connection
        cursor = self.cursor
        
        filesDict = {'//drive/folder/folder/folder/folder/SQL/SQL_Transform.sql': 'SQL TRANSFORMATIONS',
                     '//drive/folder/folder/folder/folder/SQL/SQL_Load.sql': 'SQL TABLE LOAD',
                     '//drive/folder/folder/folder/folder/SQL/SQL_Load_Incurred.sql': 'SQL TABLE LOAD'}
        
        for i in list(filesDict.keys()):
            
            # Import .sql File
            with open(i, 'r') as queries:
                sqlFile = queries.read()
        
            # all SQL commands (split on ';')
            sqlCommands = sqlFile.split(';')
        
            # Execute every command from the input file
            print(f'{filesDict[i]} STARTED AT {datetime.datetime.now()}')
            for command in sqlCommands:
                try:
                    cursor.execute(command)

                except:
                    print(f"COMMAND FAILURE: {command}")
            
            print(f'{filesDict[i]} COMPLETE AT {datetime.datetime.now()}')
            
        # CLOSE CONNECTION
        cursor.close()    