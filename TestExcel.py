#Import the following library to make use of the DispatchEx to run the macro
import os
import win32com.client as wincl

def runMacro():

    if os.path.exists("Book2.xlsm"):
        print("In Macro")
        # DispatchEx is required in the newest versions of Python.
        excel_macro = wincl.DispatchEx("Excel.application")
        excel_path = os.path.expanduser("D:\\UpworkProjects\\ExcelTest\\ExcelTest4\\Book2.xlsm")
        print(excel_path)
        workbook = excel_macro.Workbooks.Open(Filename = excel_path, ReadOnly =1)
        excel_macro.Application.Run("Book2.xlsm!Sheet1.CommandButton1_Click")
        workbook.Save()
        workbook.Close(False)
        excel_macro.Application.Quit()
        del workbook
        del excel_macro

runMacro()