Option Explicit

Dim File_obj,Read_obj,Write_obj
Dim i
Set File_obj=CreateObject("Scripting.FileSystemObject")

'MSGBOX File_obj.DriveExists("D:\") 

'MSGBOX fILE_OBJ.fOLDEReXISTS("D:\Myfolder")

' IF nOT(fILE_OBJ.fOLDEReXISTS("D:\Myfolder")) Then
	' fILE_OBJ.CreateFolder "D:\Myfolder"
' END IF

' if nOT(File_obj.FileExists("D:\Myfolder\sample.txt")) Then
	' File_obj.CreateTextFile "D:\Myfolder\sample.txt"
' end IF

' Set Write_obj=File_obj.OpenTextFile("D:\Myfolder\sample.txt",2)

' For i=65 to 90
	' Write_obj.Write(Chr(i))
' Next

' Set Read_obj=File_obj.OpenTextFile("D:\Myfolder\sample.txt",1)

' 'MSGBOX Read_obj.ReadAll
' Do Until Read_obj.AtEndOfStream

	' MSGBOX Read_obj.ReadLine
' Loop


'File_obj.CopyFolder "D:\Myfolder","E:\"

'File_obj.CopyFile "D:\Myfolder\sample.txt","E:\Myfolder\"

'File_obj.MoveFile "D:\Myfolder\sample.txt","E:\"

'File_obj.MoveFolder "D:\Myfolder","E:\"

'fILE_OBJ.DeleteFile "D:\Myfolder\sample.txt"

File_obj.DeleteFolder "D:\Myfolder"