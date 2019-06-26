import time
import sys
import os
import re
import pandas as pd
from pyxlsb import open_workbook as open_xlsb

global fileName, sheetName, fileExist, rowArray, rowNumber, helpNeeded, sqlStatement, tableName, columnNameArray, sysCliColumn, noOfRowsToCopy

def init():
	print('Script is being initialized...')
	global fileName, sheetName, fileExist, rowArray, rowNumber, helpNeeded, sqlStatement, tableName, columnNameArray, sysCliColumn, noOfRowsToCopy
	getArgs() # This will get the arguments from user input
	rowArray = []
	rowNumber = 0

	if not fileName is None and not sheetName is None and not tableName is None and not sysCliColumn is None and not noOfRowsToCopy is None:
		startScript()
	elif helpNeeded is True:
		print('')
	else:
		print('You seem to have forgotten a crucial variable when calling the script.')
		print('Please use this command to get help: python.exe export_new_format.py -h')


def startScript():
	global fileName, sheetName, fileExist, rowArray, rowNumber, helpNeeded, sqlStatement, tableName, columnNameArray, sysCliColumn, noOfRowsToCopy
	## Check if file exist
	try:
		with open(fileName) as file:
			fileExist = True
	except FileNotFoundError:
		print('File not found')
		fileExist = False

	## Save the rows in a readable array format
	## create SQL Statement from readable array
	if fileExist:
		saveExcelRowsInArray()
		createSqlStatement()

def createSqlStatement():
	global fileName, sheetName, fileExist, rowArray, rowNumber, helpNeeded, sqlStatement, tableName, columnNameArray, sysCliColumn, noOfRowsToCopy

	sqlStatement = "INSERT INTO " #um_start (column1, column2, column3, column4)"
	sqlStatement += tableName + " ("
	for columnName in columnNameArray:
		sqlStatement += columnName + ", "
	else: #For loop finished
		sqlStatement = sqlStatement[:-2]

	sqlStatement += ") VALUES " #(value1, value2, value3, value4)
	#print(rowArray)
	for row in rowArray:
		sqlStatement += "("
		for x in range(len(row)):
			#print(x)
			#print(len(rowArray))
			#print(row)
			if not row[x] is None:
				if x == int(sysCliColumn): # Sys/Cli column in Sheet
					sysCli = row[x].split('/')
					#print(sysCli)
					sys = sysCli[0].lstrip() ## This will remove whitespaces from beginning of string
					sys = sys.rstrip() ## This will remove whitespaces from end of string
					cli = sysCli[1].lstrip() ## This will remove whitespaces from beginning of string
					cli = cli.rstrip() ## This will remove whitespaces from end of string
					sqlStatement += "'" + sys + "', '" + cli + "', "
#				elif x == 2: # 2 is column C in Sheet
#					tmpValue = row[x].replace('0001', '') ## This will remove the string '0001' from value
#					tmpValue = tmpValue.rstrip() ## This will remove whitespaces from end of string
#					tmpValue = tmpValue.lstrip() ## This will remove whitespaces from beginning of string
#					sqlStatement += "'" + tmpValue + "', "
				else:
					tmpVal = row[x]
					tmpVal = tmpVal.rstrip() ## This will remove whitespaces from end of string
					tmpVal = tmpVal.lstrip() ## This will remove whitespaces from beginning of string
					sqlStatement += "'" + tmpVal + "', "
		else:
			sqlStatement = sqlStatement[:-2] ## This will remove the extra ', ' from sql string
			sqlStatement += "), "
	else:
		sqlStatement = sqlStatement[:-2] ## This will remove the extra ', ' from sql string
		sqlStatement += ";"


	print('SQL Statement: ')
	print('')
	print(sqlStatement)

## As the name says, it will go through all rows in excel sheet
## and save it into an array
def saveExcelRowsInArray():
	global fileName, sheetName, fileExist, rowArray, rowNumber, helpNeeded, sqlStatement, tableName, columnNameArray, sysCliColumn, noOfRowsToCopy

	with open_xlsb(fileName) as wb:
		with wb.get_sheet(sheetName) as sheet:
			print('Reading rows of Excel sheet')
			print('This could take a while, depending on how many rows you are going to use in the SQL Statement')
			print('')
			for row in sheet.rows(sparse=True):
				if (not rowNumber == 0):
					#print("RowNumber != 0")
					if (rowNumber <= int(noOfRowsToCopy)): # (51) This will take the first 50 rows
						## Go Through columns and put them into an array inside another array
						## This way, the data will be a lot more readable
						rowArray.append([item.v for item in row])
					else:
						break
				rowNumber = rowNumber + 1
	
	#for row in rowArray:
	#	print(row)


def getArgs():
	global fileName, sheetName, fileExist, rowArray, rowNumber, helpNeeded, sqlStatement, tableName, columnNameArray, sysCliColumn, noOfRowsToCopy
	helpNeeded = False
	columnNameArray = []
	for arg in sys.argv:
		# This is the name of the file with the user credentials
		if 'file=' in arg:
			fileName = arg.replace('file=', '') #remove 'file=' and save fileName in variable
			print('Using file:', end=" "),
			print(fileName)
		if 'sheet=' in arg:
			sheetName = arg.replace('sheet=', '')
			print('Using sheet:', end=" "),
			print(sheetName)
		if 'table=' in arg:
			tableName = arg.replace('table=', '')
			print('Using table:', end=" "),
			print(tableName)
			print('')
		if 'column' in arg:
			# Go through for loop to check what column number it is
			tmpColumn = ""
			if 'column1' in arg:
				tmpColumn = arg.replace('column1=', '')
				#columnNameArray[0] = tmpColumn
			elif 'column2' in arg:
				tmpColumn = arg.replace('column2=', '')
				#columnNameArray[1] = tmpColumn
			elif 'column3' in arg:
				tmpColumn = arg.replace('column3=', '')
				#columnNameArray[2] = tmpColumn
			elif 'column4' in arg:
				tmpColumn = arg.replace('column4=', '')
				#columnNameArray[3] = tmpColumn
			elif 'column5' in arg:
				tmpColumn = arg.replace('column5=', '')
				#columnNameArray[4] = tmpColumn
			elif 'column6' in arg:
				tmpColumn = arg.replace('column6=', '')
				#columnNameArray[5] = tmpColumn
			elif 'column7' in arg:
				tmpColumn = arg.replace('column7=', '')
				#columnNameArray[6] = tmpColumn
			elif 'column8' in arg:
				tmpColumn = arg.replace('column8=', '')
				#columnNameArray[7] = tmpColumn
			columnNameArray.append(tmpColumn)
		if 'syscli=' in arg:
			sysCliColumn = arg.replace('syscli=', '')
		if 'noOfRows=' in arg:
			noOfRowsToCopy = arg.replace('noOfRows=', '')
			
		if '-h' in arg:
			helpNeeded = True
			print('Example: python.exe export_new_format.py file=nameOfFile.xlsb sheet=Sheet1 noOfRows=noOfRowsToRead table=tableName column1=id column2=name column3=System/Client syscli=1')
			print('')
			print('You need to use the format, that was given an example of, above.')
			print('You can NOT use this script without specifying what file, sheet, database table to use. And what the names of each column in the database is. (Of cause only put the column names you want to insert data into.)')
			print('The file HAS to be with the MIME type of: .xlsb AND the file name and sheet name has to be named exactly as it is, this includes capital letters.')
			print('You HAVE to give a column number for where the Sys/Cli is stored. If System and Client is already split up into different columns in the Data Sheet, just put syscli=100')
			print('The columns should be named column1=FirstColumnInExcelSheet column2=SecondColumnInExcelSheet and so on.')
			print('IT IS VERY IMPORTANT THAT YOU SET column1, AS THE FIRST COLUMN IN THE WORKBOOK SHEET.')
			print('IF YOU DO NOT DO THIS, THE SQL STATEMENT WILL PUT THE WRONG DATA IN THE WRONG COLUMN IN THE DATABASE')
			print('ANOTHER IMPORTANT NOTE IS: The cell cannot be empty. If it is empty it will skip it, and not put it in the SQL Statement. This will result in the data also being put in the wrong column in the Database.')
			print('')
			print('Example:')
			print('If you have a Data Sheet that has these columns:')
			print('ID	Name	Sys/Cli		Object name')
			print('')
			print('And the Database table looks like this:')
			print('id	system	 client	  name	  objName')
			print('')
			print('The command should look like this:')
			print('python.exe create_sql_statement.py file=file_name.xlsb sheet=Sheet1 noOfRows=noOfRowsToRead table=DatabaseTableName column1=id column2=system column3=client column4=name column5=objName syscli=2')
			print('')
			print('Again, it does NOT matter if the column "id" is not the first column in the Database. The "objName" column does NOT have to be the fifth column in the Database either. As long as column[number] corrospond to the column number in the Data Sheet.')

	## This will set variables to None, if the user has not given anything in script call
	try:
		fileName
	except NameError:
		fileName = None
	try:
		sheetName
	except NameError:
		sheetName = None
	try:
		tableName
	except NameError:
		tableName = None
	try:
		sysCliColumn
	except NameError:
		sysCliColumn = None
	try:
		noOfRowsToCopy
	except NameError:
		noOfRowsToCopy = None


## This checks that it was the script called mainly. And calls init()
if __name__ == '__main__':
	init()