'Copy the text from the same file ("MyFirstTextFile.txt") and create an another text file 'with name "Result.txt" where you can paste the copied data


Option Explicit

Dim File_Obj,Write_Obj,Read_Obj
Dim Dest,Src,ForWrite,ForRead
Dim str

Set File_Obj=CreateObject("Scripting.FileSystemObject")

ForWrite=2
ForRead=1

Src="C:\Vrushali\Testing\MyFirstTextFile.txt"

Dest="C:\Vrushali\Result.txt"

If File_Obj.FileExists(Src) Then

	IF Not(File_Obj.FileExists(Dest)) Then
		File_Obj.CreateTextFile Dest
	End IF
	
		Set Read_Obj=File_Obj.OpenTextFile(Src,ForRead)

		Set Write_Obj=File_Obj.OpenTextFile(Dest,ForWrite)


		Do Until Read_Obj.AtEndOfStream
				str=Read_Obj.ReadLine()
				Write_Obj.WriteLine(str)
		Loop
Else
	MsgBox " File With Path "&Src&" Not Exists"
End If

set Write_Obj=Nothing
Set File_Obj=Nothing
Set Read_Obj=Nothing



