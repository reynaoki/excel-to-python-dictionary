# excel-to-python-dictionary
A Jython function that parses through an Excel file (either .xls or .xlsx) using Apache POI. Returns a dictionary where the keys are row indices and the values are a list of the contents of a single row.

## Required imports
from java.io import FileInputStream
- Make sure Jython is running

from org.apache.poi.hssf.usermodel import HSSFWorkbook
from org.apache.poi.hssf.usermodel import HSSFFormulaEvaluator
from org.apache.poi.xssf.usermodel import XSSFWorkbook
from org.apache.poi.xssf.usermodel import XSSFFormulaEvaluator
- Download 'poi-ooxml-version-yyyymmdd.jar' from http://poi.apache.org
- Fully tested on Apache POI version 3.8. Download link: https://archive.apache.org/dist/poi/release/bin/poi-bin-3.8-20120326.zip

## Function arguments
def getExcelSheetData(excelFilename, sheetName, calculateFormulas = True):

### excelFilename:
A string containing the complete filename of the Excel file that you want to extract data from. For example, "D:\sampleFile.xlsx"

### sheetName:
The name of the Excel sheet that you want to extract data from. This function can only extract data from one sheet at a time. Accepts either a string or integer (index)
Examples:
- String: "Sheet1"
- Integer: 0

### calculateFormulas:
A boolean representing whether or not the function should attempt to calculate Excel formulas. 
- True: Calculates Excel formula at runtime (e.g. if A1 contains 5 and B1 contains 10, and the formula in C1 is "=A1+B1", then the returned value for cell C1 would be 15)
- False: Returns the plaintext representation of the Excel formula (e.g. "=A1+B1")
