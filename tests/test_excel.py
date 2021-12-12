from pywinauto import Desktop,Application
import pywinauto
from app.excel import ExcelCode

excelObj=ExcelCode()
class TestExcel():

    #Open a new Excel sheet
    def test_launchExcel(self):
        try:
            excelObj.launch()
        except Exception as e:
            print("Launch of excel was not successful: "+ str(e))
            raise

    # Open a saved Excel
    def test_OpenSavedExcel(self):
        try:
            excelObj.openSavedExcel()
        except Exception as e:
            print("Opening of saved excel was not successful: "+ str(e))
            raise

    #Apply pivot table
    def test_ApplyPivot(self):
        try:
            excelObj.applyPivotTable()
        except Exception as e:
            print("Applying Pivot table was not successful: "+ str(e))
            raise

    #Save the file
    def test_SaveExcel(self):
        try:
            excelObj.saveExcel()
        except Exception as e:
            print("Save Excel was not successful: "+ str(e))
            raise