'Create .txt file at location C:\Vrushali\Testing with name "MyFirstTextFile.txt“

Option Explicit

Dim File_Obj
Dim Src

Set File_Obj=CreateObject("Scripting.FileSystemObject")


Src="C:\Vrushali\Testing\MyFirstTextFile.txt"

If Not(File_Obj.FolderExists("C:\Vrushali")) Then
	File_Obj.CreateFolder "C:\Vrushali"
End If

If Not(File_Obj.FolderExists("C:\Vrushali\Testing")) Then
	File_Obj.CreateFolder "C:\Vrushali\Testing"
End If

If Not(File_Obj.FileExists(Src)) Then
	File_Obj.CreateTextFile Src
End If

Set File_Obj=Nothing


