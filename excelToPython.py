from java.io import FileInputStream
from org.apache.poi.hssf.usermodel import HSSFWorkbook
from org.apache.poi.hssf.usermodel import HSSFFormulaEvaluator
from org.apache.poi.xssf.usermodel import XSSFWorkbook
from org.apache.poi.xssf.usermodel import XSSFFormulaEvaluator

def getExcelSheetData(excelFilename, sheetName, calculateFormulas = True):
    '''
    Parses through an Excel spreadsheet and returns a dictionary of the entire spreadsheet. 
    Keys are row numbers, and values are a list of that row's contents. Excel 
    formulas can be calculated, or displayed as the formula text. Default setting is 
    to calculate formulas.
    '''
    
    if not (excelFilename.endswith(".xls") or excelFilename.endswith(".xlsx")):
        raise Exception("Must use file extension '.xls' or '.xlsx'. This file is invalid: %s" %(excelFilename))
    
    #Open a FileInputStream (used to read contents of file as a stream of bytes)
    fis = FileInputStream(excelFilename)
    
    #Load the entire workbook
    if excelFilename.endswith(".xls"): #Excel 97-03
        workbook = HSSFWorkbook(fis)
    elif excelFilename.endswith(".xlsx"): #Excel 07 and newer
        workbook = XSSFWorkbook(fis)
    
    #Load a particular sheet (identified by string or integer)
    if isinstance(sheetName, str):
        sheet = workbook.getSheet(sheetName)
    elif isinstance(sheetName, int):
        sheet = workbook.getSheetAt(sheetName)
    else:
        raise Exception("Invalid type for sheetName. Use either str or int")

    #Get index of last row (index starts at 0)
    rows = sheet.getLastRowNum()
    
    #Get max number of columns used in spreadsheet
    cols = 0
    tmp = 0
    for r in range(0, rows + 1):
        row = sheet.getRow(r)
        if row != None:
            #getLastCellNum() returns index of last cell + 1
            tmp = sheet.getRow(r).getLastCellNum()
            #print("row: %s\t # of cells: %s" %(r,tmp))
            if tmp > cols:
                cols = tmp
    
    #Initialize sheet dictionary (key is row index, value is list of row's contents)
    sheetDict = {}
    
    #Iterate through content and put each row's contents into a list
    for r in range(0, rows + 1):
        rowList = []
        row = sheet.getRow(r)
        #Do not look at empty rows
        if row != None:
            for c in range(0,cols):
                cell = row.getCell(c)
                #Empty cells treated as empty string
                if cell is None:
                    rowList.append("")
                else:
                    #Get cell type. 0 = numeric, 1 = string, 2 = formula, 3 = blank, 4 = boolean, 5 = error
                    cellType = cell.getCellType()
                    if cellType == 2:
                        if calculateFormulas:
                            if excelFilename.endswith(".xls"):
                                formulaEvaluator = HSSFFormulaEvaluator(workbook)
                            elif excelFilename.endswith(".xlsx"):
                                formulaEvaluator = XSSFFormulaEvaluator(workbook)
                            cell = formulaEvaluator.evaluate(cell).getStringValue()
                            cell = cell.encode('UTF-8')
                        else:
                            cell = cell.toString().encode('UTF-8')
                    else:
                        cell = cell.toString().encode('UTF-8')
                    rowList.append(cell)
        #Put data for this row into the sheet dictionary
        sheetDict.update({r:rowList})
    fis.close()
    
    #For debugging, prints out the dictionary
    # for k,v in sheetDict.items():
        # print("Row %s: %s" %(k,v))
    return sheetDict
