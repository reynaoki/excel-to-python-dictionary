# excel-to-python-dictionary
A function that parses through an Excel file using Apache POI. Returns a dictionary where the keys are row indices and the values are a list of the contents of a single row.

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
