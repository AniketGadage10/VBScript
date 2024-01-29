'Create Folder at location C:\ with your name

Option Explicit

Dim File_Obj
Dim Src

Set File_Obj=CreateObject("Scripting.FileSystemObject")

Src="C:\AniketGadage"

If File_Obj.FolderExists(Src) Then
	MsgBox "Folder Is Already Exists "
Else
	File_Obj.CreateFolder Src
End If

Set File_Obj=Nothing
