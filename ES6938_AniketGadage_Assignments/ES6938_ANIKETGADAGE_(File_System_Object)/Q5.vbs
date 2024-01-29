'5. Now move newly created "Result" file at location C:\Vrushali\Testing\Result


Option Explicit

Dim File_Obj
Dim Dest,Src

Set File_Obj=CreateObject("Scripting.FileSystemObject")

Dest="C:\Vrushali\Testing\"

Src="C:\Vrushali\Result.txt"


If File_Obj.FileExists(Src) Then

	IF File_Obj.FileExists(Dest) Then
		File_Obj.MoveFile Src,Dest
	Else
		MsgBox " File With Path "&Dest&" Not Exists"
	End IF
Else
	MsgBox " File With Path "&Src&" Not Exists"
End If


Set File_Obj=Nothing





