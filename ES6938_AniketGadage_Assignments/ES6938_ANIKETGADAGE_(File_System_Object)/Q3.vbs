'Write some text under the same file - "MyFirstTextFile.txt“


Option Explicit

Dim File_Obj,Write_Obj
Dim Src,ForWrite
Dim str

ForWrite=2

Set File_Obj=CreateObject("Scripting.FileSystemObject")

Src="C:\Vrushali\Testing\MyFirstTextFile.txt"

IF File_Obj.FileExists(Src) Then
		Set Write_Obj=File_Obj.OpenTextFile(Src,ForWrite)
		str=InputBox(" Enter The Text")
		Write_Obj.WriteLine(str)
Else
	MsgBox " File With Path "&Src&" Not Exists"
End IF


set Write_Obj=Nothing
Set File_Obj=Nothing