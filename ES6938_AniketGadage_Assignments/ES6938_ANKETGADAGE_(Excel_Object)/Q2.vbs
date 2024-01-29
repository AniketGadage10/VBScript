'  Check existence of file: "MyFirstExcelFile.xlsx", If exist then open the file and enter data
'Sr.No   User ID   Password
'1       Expleo1   Password@1
'2       Expleo2   Password@2

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
	
	XLWorkSheet_Obj.Cells(1,1).Value="Sr.No"
	XLWorkSheet_Obj.Cells(1,2).Value="User ID"
	XLWorkSheet_Obj.Cells(1,3).Value="Password"
	XLWorkSheet_Obj.Cells(2,1).Value="1"
	XLWorkSheet_Obj.Cells(2,2).Value="Expleo1"
	XLWorkSheet_Obj.Cells(2,3).Value="Password@1"
	XLWorkSheet_Obj.Cells(3,1).Value="2"
	XLWorkSheet_Obj.Cells(3,2).Value="Expleo2"
	XLWorkSheet_Obj.Cells(3,3).Value="Password@2"
	
Else
	MsgBox " xlsx File With Path "&Src&"  Not Exists " 
End If


XL_Obj.Quit

Set XL_Obj=Nothing
Set XLWorkBook_Obj=Nothing
Set XLWorkSheet_Obj=Nothing