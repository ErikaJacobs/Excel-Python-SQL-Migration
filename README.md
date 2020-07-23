# Data From Hundreds of Excel Files to SQL Server
This project features the migration of hundreds of Excel files with a similar format to SQL using Python. This code also features a conversion loop to remove macros from all Excel files being migrated using VBA code, and saving a new copy of these files.

## Methods Used
* ETL

## Technologies Used
* Python
* Spyder
* SQL Server

## Packages Used
* Pandas
* Pyodbc
* Win32com
* Glob
* Numpy
* Os

# Featured Notebooks, Scripts, Analysis, or Deliverables
* [```run.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/run.py)
* [Blog Post](https://erikajacobs.netlify.app/post/how-the-itsy-bitsy-spyder-saved-my-project/)

# Other Repository Contents
* Modules
     * [```clean.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/modules/clean.py) - Basic data cleansing in Python
     * [```convert_check.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/modules/convert_check.py) - Checks if any files need to be converted/uploaded
     * [```convert_no_macro.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/modules/convert_no_macro.py) - Removes macros from Excel files and saves new copy
     * [```excel_import.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/modules/excel_import.py) - Imports Excel files to Python
     * [```main.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/modules/main.py) - Organizes execution of all modules
     * [```sql_export.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/modules/sql_export.py) - Exports data from Python to SQL Server
     * [```sql_load.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/modules/sql_load.py) - Runs code to create final tables in SQL Server
     * [```sql_transform.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/modules/sql_transform.py) - Runs code to clean data in SQL Server
     * [```warning.py```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/modules/warning.py) - Creates warning message for importing all directory Excel files
* SQL Files
     * [```SQL_Load.sql```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/SQL/SQL_Load.sql) - Creates final tables in SQL Server
     * [```SQL_Transform.sql```](https://github.com/ErikaJacobs/Excel-Python-SQL-Migration/blob/master/SQL/SQL_Transform.sql) - Cleans data in SQL Server

# Sources
* [Pyodbc](https://github.com/mkleehammer/pyodbc/wiki)

