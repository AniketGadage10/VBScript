'Create Folder at location C:\ with your name

Option Explicit

Dim File_Obj
Dim Src

Set File_Obj=CreateObject("Scripting.FileSystemObject")

Src="C:\AniketGadage"

File_Obj.CreateFolder Src
