'7. Write a program to add 5 key and value into dictionary object and
 'print all of them
 
 Option Explicit
 
	Dim Dic_Obj
	Dim skey
	
	Set Dic_Obj=CreateObject("Scripting.Dictionary")
	
	Dic_Obj.add "A","Aniket Gadage"
	Dic_Obj.add "B","Dikshant Wagh"
	Dic_Obj.add "C","Yash Bhagat"
	Dic_Obj.add "D","Ram Ghire"
	Dic_Obj.add "E","Rohit Pathak"
	
	For Each skey in Dic_Obj.Keys()
		MsgBox " Key : "&skey&" Value : "&Dic_Obj(skey)
	Next