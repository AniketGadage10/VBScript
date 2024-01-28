Option Explicit

'Dim File_Obj,Src

'Set File_Obj=CreateObject("Scripting.FileSystemObject")

'Src="E:\AniketGadage"

'Src="E:\AniketGadage\test.txt"


'File_Obj.CreateTextFile Src

'File_Obj.CreateFolder Src

'----------------------------------------------------------------------------

'Dim File_Obj,Write_Obj
'Dim ForWrite,Src,i

'Set File_Obj=CreateObject("Scripting.FileSystemObject")

'Src="E:\AniketGadage\test.txt"

'ForWrite=2

'Set Write_Obj=File_Obj.OpenTextFile(Src,ForWrite)

'For i=0 to 2
'	Write_Obj.WriteLine(InputBox("Enter The Text Massage"))

'next

'-------------------------------------------------------------------------------


'Dim File_Obj,Read_Obj
'Dim ForRead,Src,i

'Set File_Obj=CreateObject("Scripting.FileSystemObject")

'Src="E:\AniketGadage\test.txt"

'ForRead=1

'Set Read_Obj=File_Obj.OpenTextFile(Src,ForRead)

'MsgBox Read_Obj.ReadAll

'Do Until Read_Obj.AtEndOfStream
'	MsgBox Read_Obj.ReadLine()
'loop


'-------------------------------------------------------------------------------


'Dim File_Obj,Read_Obj
'Dim ForRead,Src,i

'Set File_Obj=CreateObject("Scripting.FileSystemObject")

'File_Obj.CopyFolder "E:\AniketGadage","F:\"

'File_Obj.CopyFile "E:\AniketGadage\test.txt","F:\"

'MsgBox File_Obj.DriveExists("C:")

'MsgBox File_Obj.FileExists("E:\AniketGadage\test.txt")

'MsgBox File_Obj.FolderExists("E:\AniketGadage")

'File_Obj.MoveFile "E:\AniketGadage\test.txt","F:\"

'File_Obj.MoveFolder "E:\AniketGadage","F:"

'----------------------------------------------------------------------------------

Dim Excel_Obj


Set Excel_Obj=CreateObject("Excel.Application")

Excel_Obj.ActiveWorkbook.SaveAs "E:\Aniket.xlsx"

Excel_Obj.Quit
Set Excel_Obj=Nothing






