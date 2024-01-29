'6. Now delete the text file : "MyFirstTextFile.txt“ from the location.


Option Explicit

Dim File_Obj
Dim Src

Set File_Obj=CreateObject("Scripting.FileSystemObject")

Src="C:\Vrushali\Testing\MyFirstTextFile.txt"

File_Obj.DeleteFile Src