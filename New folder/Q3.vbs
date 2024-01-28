'Write some text under the same file - "MyFirstTextFile.txt“


Option Explicit

Dim File_Obj,Write_Obj
Dim Src,ForWrite
Dim str

ForWrite=2

Set File_Obj=CreateObject("Scripting.FileSystemObject")

Src="C:\Vrushali\Testing\MyFirstTextFile.txt"

Set Write_Obj=File_Obj.OpenTextFile(Src,ForWrite)


Do 
	str=InputBox(" Enter The Text OR 'Q' For Exit")
	Write_Obj.WriteLine(str)
Loop While str<>"Q"