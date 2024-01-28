'5. Now move newly created "Result" file at location C:\Vrushali\Testing\Result


Option Explicit

Dim File_Obj
Dim Dest,Src

Set File_Obj=CreateObject("Scripting.FileSystemObject")

Dest="C:\Vrushali\Testing\"

Src="C:\Vrushali\Result.txt"

File_Obj.MoveFile Src,Dest