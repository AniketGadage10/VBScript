'Input : sString = Array("I","am","learning","Vbscript")
'Write a program to get below output
'output : "I am learning Vbscript"



' sString = Array("I","am","learning","Vbscript")

' Wscript.Echo Join(sString," ")

'------------------------------------------------------------------

Option Explicit

Dim Input_Arr()
Dim Size_1,Int_I
Size_1=Int(InputBox("Enter The Size Of 1st Array"))

Redim Input_Arr(Size_1)

For Int_I=0 to Size_1

	Input_Arr(Int_I)=InputBox("Enter The "&Int_I&"TH Element Of Array")
	
next

Wscript.Echo Join(Input_Arr," ")
