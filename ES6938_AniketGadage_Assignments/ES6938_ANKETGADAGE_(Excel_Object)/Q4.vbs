'Now add one mores sheet under the same excel with name "master"


Option Explicit

Dim XL_Obj,XLWorkBook_Obj,XLWorkSheet_Obj
Dim xl_Src

Set XL_Obj=CreateObject("Excel.Application")

xl_Src="C:\MyFirstExcelFile.xlsx"

If XL_Obj.fileExists(Src) then

	Set XLWorkBook_Obj=XL_Obj.WorkBooks.open(Src)
	Set XLWorkSheet_Obj=XLWorkBook_Obj.Worksheet.add
	XLWorkSheet_Obj.name="master"
Else
	MsgBox " xlsx File With Path "&Src&"  Not Exists " 
End If

XL_Obj.Quit

Set XL_Obj=Nothing
Set XLWorkBook_Obj=Nothing
Set XLWorkSheet_Obj=Nothing