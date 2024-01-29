'6. Now try to delete 2nd row data from the first sheet.

Option Explicit

Dim XL_Obj,XLWorkBook_Obj,XLWorkSheet_Obj
Dim Src
Dim irow_count,icol_count,irow,icol

Set XL_Obj=CreateObject("Excel.Application")

Src="C:\MyFirstExcelFile.xlsx"

XL_Obj.Visible=True

If XL_Obj.fileExists(Src) then

	Set XLWorkBook_Obj=XL_Obj.WorkBooks.open(Src)
	Set XLWorkSheet_Obj=XLWorkBook_Obj.Worksheet(1)
	
	XLWorkSheet_Obj.Rows(2).delete
	
Else
	MsgBox " xlsx File With Path "&Src&"  Not Exists " 
End If


XL_Obj.Quit

Set XL_Obj=Nothing
Set XLWorkBook_Obj=Nothing
Set XLWorkSheet_Obj=Nothing