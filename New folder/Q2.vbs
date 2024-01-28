'Create .txt file at location C:\Vrushali\Testing with name "MyFirstTextFile.txt“

Option Explicit

Dim File_Obj
Dim Src

Set File_Obj=CreateObject("Scripting.FileSystemObject")

Src="C:\Vrushali\Testing\MyFirstTextFile.txt"

File_Obj.CreateTextFile Src