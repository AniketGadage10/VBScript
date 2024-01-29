'Now fetch the data whose Sr. No is 2 and wright them in text file "Result.txt"

Option Explicit

Dim XL_Obj,XLWorkBook_Obj,XLWorkSheet_Obj
Dim File_Obj,Write_Obj
Dim xl_Src,file_Src,str
Dim Int_col,forwrite

Set XL_Obj=CreateObject("Excel.Application")
Set File_Obj=CreateObject("Scripting.FileSystemObject")
forwrite=2
xl_Src="C:\MyFirstExcelFile.xlsx"
file_Src="C:\Result.txt"

If XL_Obj.fileExists(Src) then

	Set XLWorkBook_Obj=XL_Obj.WorkBooks.open(Src)
	Set XLWorkSheet_Obj=XLWorkBook_Obj.Worksheet(1)
	
	IF Not(File_Obj.FileExists(file_Src)) then
		File_Obj.CreateTextFile file_Src
	End IF
	Set Write_Obj=File_Obj.OpenTextFile(Src,forwrite)
	
	For Int_col=1 to 3 
		str=XLWorkSheet_Obj.Cells(1,Int_col).Value&" = "&XLWorkSheet_Obj.Cells(3,Int_col).Value
		Write_Obj.WriteLine(str)
	Next
	
Else
	MsgBox " xlsx File With Path "&Src&"  Not Exists " 
End If

XL_Obj.Quit

Set XL_Obj=Nothing
Set XLWorkBook_Obj=Nothing
Set XLWorkSheet_Obj=Nothing