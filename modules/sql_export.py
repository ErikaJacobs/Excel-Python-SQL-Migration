# Exports Excel data to SQL Server

def SQL_Exp_All(df):
    import pyodbc
    import datetime
    import pandas as pd
    
    #Creates Connection
    conn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};'
                          'Server=Server Name'
                          'Database=Database Name;'
                          'Trusted_Connection=yes;')
    cursor = conn.cursor()
    
    # Create Table in SQL
    cursor.execute('USE Database Name;' 
                   "IF OBJECT_ID('???.TableName', 'U') IS NOT NULL DROP TABLE ???.TableName;"
                   """CREATE TABLE ???.TableName (variable1 NVARCHAR(30), variable2 NVARCHAR(30), variable3 NVARCHAR(30), \
                   variable4 NVARCHAR(30), variable5 NVARCHAR(30), variable6 NVARCHAR(30), variable7 NVARCHAR(30), variable8 NVARCHAR(30), \
                   variable9 NVARCHAR(30), variable10 NVARCHAR(30), variable11 NVARCHAR(30), variable12 NVARCHAR(30), \
                   variable13 DECIMAL (32,6), variable14 DECIMAL (32,6), variable15 DECIMAL (32,6), variable16 DECIMAL (32,6), \
                   variable17 DECIMAL (32,6), variable18 DECIMAL (32,6), variable19 DECIMAL (32,6), variable20 DECIMAL (32,6), \
                   variable21 DECIMAL (32,6), variable22 DECIMAL (32,6), variable23 DECIMAL (32,6), variable24 DECIMAL (32,6), \
                   variable25 DECIMAL (32,6), variable26 DECIMAL (32,6), variable27 DECIMAL (32,6), variable28 DECIMAL (32,6), \
                   variable29 DECIMAL (32,6), variable30 DECIMAL (32,6), variable31 DECIMAL (32,6), \
                   variable32 DECIMAL (32,6), variable33 DECIMAL (32,6), variable34 DECIMAL (32,6), variable35 DECIMAL (32,6), \
                   variable36 DECIMAL (32,6), variable37 DECIMAL (32,6), variable38 DECIMAL (32,6), \
                   variable39 DECIMAL (32,6), variable40 DECIMAL (32,6), variable41 DECIMAL (32,6), \
                   variable42 DECIMAL (32,6), variable43 DECIMAL (32,6), variable44 DECIMAL (32,6), \
                   variable45 DECIMAL (32,6), variable46 DECIMAL (32,6), \
                   variable47 DECIMAL (32,6), variable48 DECIMAL (32,6), variable49 DECIMAL (32,6), \
                   variable50 DECIMAL (32,6), variable51 DECIMAL (32,6), variable52 DECIMAL (32,6), \
                   variable53 DECIMAL (32,6), variable54 DECIMAL (32,6), variable55 DECIMAL (32,6), \
                   variable56 DECIMAL (32,6), variable57 DECIMAL (32,6), variable58 DECIMAL (32,6), \
                   variable59 DECIMAL (32,6), variable60 DECIMAL (32,6), variable61 DECIMAL (32,6), \
                   variable62 DECIMAL (32,6), variable63 DECIMAL (32,6), variable64 DECIMAL (32,6), \
                   variable65 DECIMAL (32,6), variable66 DECIMAL (32,6), variable67 DECIMAL (32,6), \
                   variable68 DECIMAL (32,6), variable69 DECIMAL (32,6), variable70 DECIMAL (32,6), \
                   variable71 DECIMAL (32,6), variable72 DECIMAL (32,6), variable73 DECIMAL (32,6), \
                   variable74 DECIMAL (32,6), variable75 DECIMAL (32,6), variable76 DECIMAL (32,6), variable77 DECIMAL (32,6), \
                   variable78 DECIMAL (32,6), variable79 DECIMAL (32,6), variable80 DECIMAL (32,6), variable81 DECIMAL (32,6), \
                   variable82 DECIMAL (32,6), variable83 DECIMAL (32,6), variable84 DECIMAL (32,6), variable85 DECIMAL (32,6), \
                   variable86 DECIMAL (32,6), variable87 DECIMAL (32,6), variable88 DECIMAL (32,6), \
                   variable89 DECIMAL (32,6), variable90 DECIMAL (32,6), variable91 DECIMAL (32,6), variable92 DECIMAL (32,6), variable93 DECIMAL (32,6), \
                   variable94 DECIMAL (32,6), variable95 DECIMAL (32,6), variable96 DECIMAL (32,6), variable97 DECIMAL (32,6), \
                   variable98 DECIMAL (32,6), variable99 DECIMAL (32,6), variable100 NVARCHAR(30), variable101 DECIMAL (32,6), \
                   variable102 DECIMAL (32,6), variable103 DECIMAL (32,6), variable104 DECIMAL (32,6), variable105 DECIMAL (32,6), \
                   variable106 DECIMAL (32,6), variable107 DECIMAL (32,6), variable108 DECIMAL (32,6), variable109 DECIMAL (32,6), \
                   variable110 DECIMAL (32,6), variable111 DECIMAL (32,6), variable112 DECIMAL (32,6), variable113 DECIMAL (32,6), \
                   variable114 DECIMAL (32,6), variable115 DECIMAL (32,6), variable116 DECIMAL (32,6), variable117 DECIMAL (32,6), \
                   variable118 DECIMAL (32,6), variable119 DECIMAL (32,6), variable120 DECIMAL (32,6), variable121 DECIMAL (32,6), \
                   variable122 DECIMAL (32,6), variable123 DECIMAL (32,6), variable124 DECIMAL (32,6), variable125 DECIMAL (32,6), \
                   variable126 DECIMAL (32,6), variable127 DECIMAL (32,6), variable128 DECIMAL (32,6), variable129 DECIMAL (32,6), \
                   variable130 DECIMAL (32,6), variable131 DECIMAL (32,6), variable132 DECIMAL (32,6), variable133 DECIMAL (32,6), \
                   variable134 DECIMAL (32,6), variable135 DECIMAL (32,6), variable136 DECIMAL (32,6), \
                   variable137 DECIMAL (32,6), variable138 DECIMAL (32,6), variable139 DECIMAL (32,6), \
                   variable140 DECIMAL (32,6), variable141 DECIMAL (32,6), variable142 DECIMAL (32,6), \
                   variable143 DECIMAL (32,6), variable144 DECIMAL (32,6), variable145 DECIMAL (32,6), \
                   variable146 DECIMAL (32,6), variable147 DECIMAL (32,6), variable148 DECIMAL (32,6), \
                   variable149 DECIMAL (32,6), variable150 DECIMAL (32,6), \
                   variable151 DECIMAL (32,6), variable152 NVARCHAR (30), variable153 NVARCHAR (30), \
                   variable154 NVARCHAR (30), variable155 DECIMAL (32,6), variable156 DECIMAL (32,6))""")
    
    conn.commit()
    
    #Inserting the data into SQL from Python
    
    print('SQL EXPORT PROCESS STARTED AT {}'.format(datetime.datetime.now()))
    
    for index,row in df.iterrows():
        cursor.execute("""INSERT INTO ???.TableName VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?)""", (row['variable1'], row['variable2'], row['variable3'], \
                       row['variable4'], row['variable5'], row['variable6'], row['variable7'], row['variable8'], row['variable9'], \
                       row['variable10'], row['variable11'], row['variable12'], row['variable13'], row['variable14'], row['variable15'], \
                       row['variable16'], row['variable17'], row['variable18'], row['variable19'], row['variable20'], row['variable21'], \
                       row['variable22'], row['variable23'], row['variable24'], row['variable25'], row['variable26'], \
                       row['variable27'], row['variable28'], row['variable29'], row['variable30'], \
                       row['variable31'], row['variable32'], row['variable33'], row['variable34'], \
                       row['variable35'], row['variable36'], row['variable37'], row['variable38'], \
                       row['variable39'], row['variable40'], row['variable41'], row['variable42'], \
                       row['variable43'], row['variable44'], row['variable45'], \
                       row['variable46'], row['variable47'], row['variable48'], \
                       row['variable49'], row['variable50'], row['variable51'], row['variable52'], \
                       row['variable53'], row['variable54'], row['variable55'], \
                       row['variable56'], row['variable57'], row['variable58'], row['variable59'], \
                       row['variable60'], row['variable61'], row['variable62'], row['variable63'], \
                       row['variable64'], row['variable65'], row['variable66'], \
                       row['variable67'], row['variable68'], row['variable69'], \
                       row['variable70'], row['variable71'], row['variable72'], \
                       row['variable73'], row['variable74'], row['variable75'], row['variable76'], \
                       row['variable77'], row['variable78'], row['variable79'], row['variable80'], \
                       row['variable81'], row['variable82'], row['variable83'], row['variable84'], \
                       row['variable85'], row['variable86'], row['variable87'], row['variable88'], \
                       row['variable89'], row['variable90'], row['variable91'], row['variable92'], row['variable93'], row['variable94'], \
                       row['variable95'], row['variable96'], row['variable97'], row['variable98'], row['variable99'], \
                       row['variable100'], row['variable101'], row['variable102'], row['variable103'], row['variable104'], \
                       row['variable105'], row['variable106'], row['variable107'], row['variable108'], row['variable109'], row['variable110'], \
                       row['variable111'], row['variable112'], row['variable113'], row['variable114'], row['variable115'], row['variable116'], \
                       row['variable117'], row['variable118'], row['variable119'], row['variable120'], row['variable121'], \
                       row['variable122'], row['variable123'], row['variable124'], row['variable125'], row['variable126'], \
                       row['variable127'], row['variable128'], row['variable129'], row['variable130'], row['variable131'], \
                       row['variable132'], row['variable133'], row['variable134'], row['variable135'], \
                       row['variable136'], row['variable137'], row['variable138'], row['variable139'], \
                       row['variable140'], row['variable141'], row['variable142'], row['variable143'], \
                       row['variable144'], row['variable145'], row['variable146'], row['variable147'], \
                       row['variable148'], row['variable149'], row['variable150'], \
                       row['variable151'], row['variable152'], row['variable153'], row['variable154'], \
                       row['variable155'], row['variable156')))
        conn.commit()
    
    conn.close()
    
    # UPDATE BACK-UP TABLE
    
    cursor.execute("IF OBJECT_ID('???.TableName_Backup', 'U') IS NOT NULL DROP TABLE ???.TableName_Backup;")
    
    conn.commit()
    
    cursor.execute('''SELECT * \
                   INTO ???.TableName_Backup \
                   FROM ???.TableName''')
    conn.commit()
    
    print('SQL EXPORT PROCESS COMPLETED AT {}'.format(datetime.datetime.now()))
    
#%%
    
def SQL_Exp_Some(df):
    import pyodbc
    import datetime
    import pandas as pd

    #Creates Connection
    conn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};'
                          'Server=Server Name'
                          'Database=Database Name;'
                          'Trusted_Connection=yes;')
    cursor = conn.cursor()
    
    # Create Table in SQL
    cursor.execute('USE Database Name;' 
                   """CREATE TABLE ???.#PLTemp (variable1 NVARCHAR(30), variable2 NVARCHAR(30), variable3 NVARCHAR(30), \
                   variable4 NVARCHAR(30), variable5 NVARCHAR(30), variable6 NVARCHAR(30), variable7 NVARCHAR(30), variable8 NVARCHAR(30), \
                   variable9 NVARCHAR(30), variable10 NVARCHAR(30), variable11 NVARCHAR(30), variable12 NVARCHAR(30), \
                   variable13 DECIMAL (32,6), variable14 DECIMAL (32,6), variable15 DECIMAL (32,6), variable16 DECIMAL (32,6), \
                   variable17 DECIMAL (32,6), variable18 DECIMAL (32,6), variable19 DECIMAL (32,6), variable20 DECIMAL (32,6), \
                   variable21 DECIMAL (32,6), variable22 DECIMAL (32,6), variable23 DECIMAL (32,6), variable24 DECIMAL (32,6), \
                   variable25 DECIMAL (32,6), variable26 DECIMAL (32,6), variable27 DECIMAL (32,6), variable28 DECIMAL (32,6), \
                   variable29 DECIMAL (32,6), variable30 DECIMAL (32,6), variable31 DECIMAL (32,6), \
                   variable32 DECIMAL (32,6), variable33 DECIMAL (32,6), variable34 DECIMAL (32,6), variable35 DECIMAL (32,6), \
                   variable36 DECIMAL (32,6), variable37 DECIMAL (32,6), variable38 DECIMAL (32,6), \
                   variable39 DECIMAL (32,6), variable40 DECIMAL (32,6), variable41 DECIMAL (32,6), \
                   variable42 DECIMAL (32,6), variable43 DECIMAL (32,6), variable44 DECIMAL (32,6), \
                   variable45 DECIMAL (32,6), variable46 DECIMAL (32,6), \
                   variable47 DECIMAL (32,6), variable48 DECIMAL (32,6), variable49 DECIMAL (32,6), \
                   variable50 DECIMAL (32,6), variable51 DECIMAL (32,6), variable52 DECIMAL (32,6), \
                   variable53 DECIMAL (32,6), variable54 DECIMAL (32,6), variable55 DECIMAL (32,6), \
                   variable56 DECIMAL (32,6), variable57 DECIMAL (32,6), variable58 DECIMAL (32,6), \
                   variable59 DECIMAL (32,6), variable60 DECIMAL (32,6), variable61 DECIMAL (32,6), \
                   variable62 DECIMAL (32,6), variable63 DECIMAL (32,6), variable64 DECIMAL (32,6), \
                   variable65 DECIMAL (32,6), variable66 DECIMAL (32,6), variable67 DECIMAL (32,6), \
                   variable68 DECIMAL (32,6), variable69 DECIMAL (32,6), variable70 DECIMAL (32,6), \
                   variable71 DECIMAL (32,6), variable72 DECIMAL (32,6), variable73 DECIMAL (32,6), \
                   variable74 DECIMAL (32,6), variable75 DECIMAL (32,6), variable76 DECIMAL (32,6), variable77 DECIMAL (32,6), \
                   variable78 DECIMAL (32,6), variable79 DECIMAL (32,6), variable80 DECIMAL (32,6), variable81 DECIMAL (32,6), \
                   variable82 DECIMAL (32,6), variable83 DECIMAL (32,6), variable84 DECIMAL (32,6), variable85 DECIMAL (32,6), \
                   variable86 DECIMAL (32,6), variable87 DECIMAL (32,6), variable88 DECIMAL (32,6), \
                   variable89 DECIMAL (32,6), variable90 DECIMAL (32,6), variable91 DECIMAL (32,6), variable92 DECIMAL (32,6), variable93 DECIMAL (32,6), \
                   variable94 DECIMAL (32,6), variable95 DECIMAL (32,6), variable96 DECIMAL (32,6), variable97 DECIMAL (32,6), \
                   variable98 DECIMAL (32,6), variable99 DECIMAL (32,6), variable100 NVARCHAR(30), variable101 DECIMAL (32,6), \
                   variable102 DECIMAL (32,6), variable103 DECIMAL (32,6), variable104 DECIMAL (32,6), variable105 DECIMAL (32,6), \
                   variable106 DECIMAL (32,6), variable107 DECIMAL (32,6), variable108 DECIMAL (32,6), variable109 DECIMAL (32,6), \
                   variable110 DECIMAL (32,6), variable111 DECIMAL (32,6), variable112 DECIMAL (32,6), variable113 DECIMAL (32,6), \
                   variable114 DECIMAL (32,6), variable115 DECIMAL (32,6), variable116 DECIMAL (32,6), variable117 DECIMAL (32,6), \
                   variable118 DECIMAL (32,6), variable119 DECIMAL (32,6), variable120 DECIMAL (32,6), variable121 DECIMAL (32,6), \
                   variable122 DECIMAL (32,6), variable123 DECIMAL (32,6), variable124 DECIMAL (32,6), variable125 DECIMAL (32,6), \
                   variable126 DECIMAL (32,6), variable127 DECIMAL (32,6), variable128 DECIMAL (32,6), variable129 DECIMAL (32,6), \
                   variable130 DECIMAL (32,6), variable131 DECIMAL (32,6), variable132 DECIMAL (32,6), variable133 DECIMAL (32,6), \
                   variable134 DECIMAL (32,6), variable135 DECIMAL (32,6), variable136 DECIMAL (32,6), \
                   variable137 DECIMAL (32,6), variable138 DECIMAL (32,6), variable139 DECIMAL (32,6), \
                   variable140 DECIMAL (32,6), variable141 DECIMAL (32,6), variable142 DECIMAL (32,6), \
                   variable143 DECIMAL (32,6), variable144 DECIMAL (32,6), variable145 DECIMAL (32,6), \
                   variable146 DECIMAL (32,6), variable147 DECIMAL (32,6), variable148 DECIMAL (32,6), \
                   variable149 DECIMAL (32,6), variable150 DECIMAL (32,6), \
                   variable151 DECIMAL (32,6), variable152 NVARCHAR (30), variable153 NVARCHAR (30), \
                   variable154 NVARCHAR (30), variable155 DECIMAL (32,6), variable156 DECIMAL (32,6))""")
    
    conn.commit()
    
    #Inserting the data into temp SQL table from Python
    
    print('SQL EXPORT PROCESS STARTED AT {}'.format(datetime.datetime.now()))
    
    for index,row in df.iterrows():
        cursor.execute("""INSERT INTO ???.#PLTemp VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, \
                                                             ?, ?)""", (row['variable1'], row['variable2'], row['variable3'], \
                       row['variable4'], row['variable5'], row['variable6'], row['variable7'], row['variable8'], row['variable9'], \
                       row['variable10'], row['variable11'], row['variable12'], row['variable13'], row['variable14'], row['variable15'], \
                       row['variable16'], row['variable17'], row['variable18'], row['variable19'], row['variable20'], row['variable21'], \
                       row['variable22'], row['variable23'], row['variable24'], row['variable25'], row['variable26'], \
                       row['variable27'], row['variable28'], row['variable29'], row['variable30'], \
                       row['variable31'], row['variable32'], row['variable33'], row['variable34'], \
                       row['variable35'], row['variable36'], row['variable37'], row['variable38'], \
                       row['variable39'], row['variable40'], row['variable41'], row['variable42'], \
                       row['variable43'], row['variable44'], row['variable45'], \
                       row['variable46'], row['variable47'], row['variable48'], \
                       row['variable49'], row['variable50'], row['variable51'], row['variable52'], \
                       row['variable53'], row['variable54'], row['variable55'], \
                       row['variable56'], row['variable57'], row['variable58'], row['variable59'], \
                       row['variable60'], row['variable61'], row['variable62'], row['variable63'], \
                       row['variable64'], row['variable65'], row['variable66'], \
                       row['variable67'], row['variable68'], row['variable69'], \
                       row['variable70'], row['variable71'], row['variable72'], \
                       row['variable73'], row['variable74'], row['variable75'], row['variable76'], \
                       row['variable77'], row['variable78'], row['variable79'], row['variable80'], \
                       row['variable81'], row['variable82'], row['variable83'], row['variable84'], \
                       row['variable85'], row['variable86'], row['variable87'], row['variable88'], \
                       row['variable89'], row['variable90'], row['variable91'], row['variable92'], row['variable93'], row['variable94'], \
                       row['variable95'], row['variable96'], row['variable97'], row['variable98'], row['variable99'], \
                       row['variable100'], row['variable101'], row['variable102'], row['variable103'], row['variable104'], \
                       row['variable105'], row['variable106'], row['variable107'], row['variable108'], row['variable109'], row['variable110'], \
                       row['variable111'], row['variable112'], row['variable113'], row['variable114'], row['variable115'], row['variable116'], \
                       row['variable117'], row['variable118'], row['variable119'], row['variable120'], row['variable121'], \
                       row['variable122'], row['variable123'], row['variable124'], row['variable125'], row['variable126'], \
                       row['variable127'], row['variable128'], row['variable129'], row['variable130'], row['variable131'], \
                       row['variable132'], row['variable133'], row['variable134'], row['variable135'], \
                       row['variable136'], row['variable137'], row['variable138'], row['variable139'], \
                       row['variable140'], row['variable141'], row['variable142'], row['variable143'], \
                       row['variable144'], row['variable145'], row['variable146'], row['variable147'], \
                       row['variable148'], row['variable149'], row['variable150'], \
                       row['variable151'], row['variable152'], row['variable153'], row['variable154'], \
                       row['variable155'], row['variable156')))
        conn.commit()
    
    print('SQL EXPORT PROCESS COMPLETED AT {}'.format(datetime.datetime.now()))
    
    # DELETE OLD RECORDS 
    cursor.execute('''DELETE FROM ???.TableName \
                   WHERE variable154 in(SELECT DISTINCT variable154 FROM ???.#PLTemp) \
                   and variable9 in(SELECT DISTINCT variable9 FROM ???.#PLTemp)''')
    conn.commit()
    
    # CREATE UNION BETWEEN HISTORY AND NEW RECORDS
    cursor.execute('''SELECT * \
                   INTO ???.#PLTemp3 \
                   FROM ???.TableName \
                   UNION ALL \
                   SELECT * \
                   FROM ???.#PLTemp''')
    conn.commit()
    
    
    # UPDATE TABLE WITH UNION
    cursor.execute('''DROP TABLE ???.TableName \
                   SELECT * \
                   INTO ???.TableName \
                   FROM ???.#PLTemp3''')
    conn.commit()
    
    # UPDATE BACK-UP TABLE
    
    cursor.execute("IF OBJECT_ID('???.TableName_Backup', 'U') IS NOT NULL DROP TABLE ???.TableName_Backup;")
    
    conn.commit()
    
    cursor.execute('''SELECT * \
                   INTO ???.TableName_Backup \
                   FROM ???.TableName''')
    conn.commit()
    
    # CLOSE CONNECTION
    
    conn.close()
    
#%%