#
# Serializes various Excel loaders for simplicity of use
# Daniel "Shadowbranch" Moree
# Python 3

import xlwt, xlutils, openpyxl, xlrd
## Loads a given Excel file
#  
#  @param fileName	string	Full filename and path
def loadExcelFile(fileName):
	try:
		excelFile = openpyxl.load_workbook(fileName)
		return excelFile
	except:
		try:
			excelFile = xlrd.open_workbook(fileName)
			return excelFile
		except:
			print("Error: Unable to open file")
	
## Returns the sheets from an Excel File
#  
#  @param excelFile	object	A loaded excel workbook object
def getSheetNames(excelFile):
	try:
		sheets = excelFile.get_sheet_names()
		return sheets
	except:
		try:
			sheets = excelFile.sheet_names()
			return sheets
		except:
			print("Error: Unable get sheet names")

## Returns the based sheet as an object
#  
#  @param excelFile	object	A loaded excel workbook object
#  @param sheetName	string	Name of the sheet to return
def getSheet(excelFile, sheetName):
	try:
		sheet = excelFile[sheetName]
		return sheet
	except:
		try:
			sheet = excelFile.sheet_by_name(sheetName)
			return sheet
		except:
			print("Error: Unable get sheet")

## Returns cell value
#  
#  @param sheet	object	Sheet to pull information from
#  @param col	int	or string	Column number or letter set
#  @param row	int		Row number
def getCellValue(sheet, col, row):
	try:
		if(str(col).isdigit() == True):
			col = int(col)
			alphaList = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
			mod = col % len(alphaList)
			precursor = ""
			if(col >= len(alphaList)):
				pos = int(col / len(alphaList)) -1
				precursor = alphaList[pos]
			value = sheet[precursor+alphaList[mod]+str(row)].value
		else:
			value = sheet[str(col)+str(row)].value
		return value
	except:
		try:
			value = sheet.cell_value(int(row)-1, int(col))
			return value
		except IndexError:
			print("Index out of range")
