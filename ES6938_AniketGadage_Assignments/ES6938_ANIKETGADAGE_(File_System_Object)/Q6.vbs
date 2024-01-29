'6. Now delete the text file : "MyFirstTextFile.txt“ from the location.


Option Explicit

Dim File_Obj
Dim Src

Set File_Obj=CreateObject("Scripting.FileSystemObject")

Src="C:\Vrushali\Testing\MyFirstTextFile.txt"

IF File_Obj.FileExists(Src) Then
	File_Obj.DeleteFile Src
Else
	MsgBox " File With Path "&Src&" Not Exists"
End IF

Set File_Obj=Nothing