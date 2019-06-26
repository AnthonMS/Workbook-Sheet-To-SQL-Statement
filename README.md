<h1> Documentation for Python script to create SQL Statement from .xlsb file </h1>

To run the script you need python, pip, pandas and pyxlsb 
Install Python as you would normally. 

Test if pip was installed with python, by running this command in a command prompt: `pip --version`

If pip was not recognized or you do not get a return string looking something like this: 
"pip 18.1 from c:\users\asteiness\appdata\local\programs\python\python37-32\lib\site-packages\pip (python 3.7)" 
Then download pip python installer from this link: https://bootstrap.pypa.io/get-pip.py
After this python executable has been securely downloaded, run the command: `python.exe get-pip.py`

When pip has been installed, you will need to install some python dependencies. This is done by running these commands: `pip install pandas` & `pip install pyxlsb`

When everything that is needed, has been installed, you are ready to start using the python script to create SQL statements from .xlsb files.

Put the script in a folder that is easy to find from command prompt, and place the .xlsb file inside this folder as well.
For easier use, rename the .xlsb file, to something that is not to long. (IMPORTANT: the file cannot hold spaces, as the script does not support that.

When the create_sql_statement.py and the .xlsb file is in the same folder, you can run this command to get help: `python.exe create_sql_statement.py -h`

This will show a quick help menu, nothing that is not already documented here.

An example of running the create_sql_statement.py:

    python.exe create_sql_statement.py file=file_name.xlsb sheet=Sheet1 noOfRows=NoOfRowsToReadFromSheet table=DatabaseTableName column1=DbColumn1 column2=DbColumn2 column3=DbColumn3 column4=DbColumn4 syscli=Sys/CliColumnInSheet

The column names has to be the same as the names in the database. So if the database has the columns: id, name, phone. You will put the column1=id column2=name column3=phone. UNLESS phone is column 2/B and name is 3/C in sheet, you will put column1=id column2=phone column3=name

You can add columns from column1 to and including column8. To use more columns, you will need to add more in the 'getArgs()' function inside the script. You do not have to use all 8 columns, they are just there for good measure.

IMPORTANT: column1 has to be the first column of the Excel Sheet/Workbook Sheet. So if column 1 in the sheet is equal to the ID of the Database table, you will HAVE to put column1=ID. ID does NOT have to be the first column in the database, it could be the last for all it cares.

