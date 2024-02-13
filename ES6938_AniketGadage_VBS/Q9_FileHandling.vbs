'9. Write a program to write your 3 line address in text file


Option Explicit

	Dim File_obj,Write_obj
	Dim iWrite
	Dim sCompany,sLocation,sCity
	
	sCompany="Expleo Solution Private Limited "
	sLocation= "Hinjewadi Phase 3,Hinjewadi Rajiv Gandhi Infotech Park"
	sCity="Pune, Maharashtra 411057"
	iWrite=8
	
	Set File_obj=CreateObject("Scripting.FileSystemObject")
	
	If File_Obj.DriveExists("C") Then
		If Not( File_obj.FolderExists("C:\Myfolder")) Then
			File_obj.CreateFolder "C:\Myfolder"
			end if
		If Not(File_obj.FileExists("C:\Myfolder\write.txt")) Then
			File_obj.CreateTextFile "C:\Myfolder\write.txt"
		end if

	Set Write_obj=File_obj.OpenTextFile("C:\Myfolder\write.txt",iWrite)
		
		Write_obj.WriteLine(sCompany)
		Write_obj.WriteLine(sLocation)
		Write_obj.WriteLine(sCity)

	END if

	Write_obj.Close
	Set File_Obj=Nothing
	Set Write_obj=Nothing






























'fILE_OBJ.cOPYfOLDER "C:\copy","C:\Users\agadage\Desktop\"

'File_obj.cOPYfILE "C:\copy\write.txt","C:\Users\agadage\Desktop\"

'fILE_OBJ.MOVEfOLDER "C:\copy","C:\Users\agadage\Desktop\"

'fILE_OBJ.mOVEfILE "C:\MyFolder\MyDetails.txt","C:\Users\agadage\Desktop\"

'fiLE_OBJ.DeleteFolder "C:\Users\agadage\Desktop\NEW"

'fiLE_OBJ.dELETEfILE "C:\Users\agadage\Desktop\MyDetails.txt"

' Msgbox File_Obj.DriveExists("C")

' if Not( File_obj.FolderExists("C:\copy")) Then
' File_obj.CreateFolder "C:\copy"
' end if


' If File_obj.FileExists("C:\copy\write.txt") Then
' File_obj.CreateTextFile "C:\copy\write.txt"
' end if


' Set Write_obj=File_obj.OpenTextFile("C:\copy\write.txt",2)

	' For i =65 to 90 Step 1
		' Write_obj.WriteLine(Chr(i))
	' Next
	
' Set Read_obj=File_obj.OpenTextFile("C:\copy\write.txt",1)

' 'Msgbox Read_obj.ReadAll

' do until Read_obj.AtEndOfStream

	' msgbox Read_obj.ReadLine

' Loop