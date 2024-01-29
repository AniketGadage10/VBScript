' Create an excel file: "MyFirstExcelFile.xlsx"

Option Explicit

Dim XL_Obj,XLWorkBook_Obj
Dim Src

Set XL_Obj=CreateObject("Excel.Application")

Src="C:\MyFirstExcelFile.xlsx"

XL_Obj.WorkBooks.Add

XL_Obj.ActiveWork.saveas Src

XL_Obj.Quit

Set XL_Obj=Nothing
