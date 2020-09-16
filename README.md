# Data From Hundreds of Excel Files to SQL Server
This project features the migration of hundreds of Excel files with a similar format to SQL using Python. This code also features a conversion loop to remove macros from all Excel files being migrated using VBA code, and saving a new copy of these files.

## Methods Used
* ETL

## Technologies Used
* Python
* Spyder
* SQL Server

## Environment
* win32com
* glob
* os
* sys
* datetime
* numpy - 1.16.5
* pandas - 0.25.1
* sqlalchemy - 1.3.9

# Featured Notebooks, Scripts, Analysis, or Deliverables
* [```run.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/run.py)


# Other Repository Contents
* Modules
    * [```convert_import.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/modules/convert_import.py)
        * Checks if any files need to be converted/uploaded
        * Removes macros from Excel files and saves new converted copy
        * Imports Excel files to Python
        * Basic data cleansing in Python
    * [```main.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/modules/main.py) - Organizes execution of all modules
    * [```sql.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/modules/sql.py)
        * Exports data from Python to SQL Server
        * Runs code to clean data in SQL Server
        * Runs code to create final tables in SQL
    * [```warning.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/modules/warning.py) - Creates warning message for importing all directory Excel files

* SQL Files
    * [```SQL_Load.sql```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/SQL/SQL_Load.sql) - Creates final tables in SQL Server
    * [```SQL_Transform.sql```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/SQL/SQL_Transform.sql) - Cleans data in SQL Server
