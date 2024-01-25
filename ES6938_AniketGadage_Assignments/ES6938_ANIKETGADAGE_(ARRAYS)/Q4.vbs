'input = str = "Apple~Grapes~Banana~Guava~Blueberries"
'Write a program which gives one array containg fruits name starting with "G"

Option Explicit

Dim input_Str
Dim arrFruits()
Dim Filter_arrFruits()
Dim index,int_i
Dim Comp_Char
index=0
input_Str = "Apple~Grapes~Banana~Guava~Blueberries"

arrFruits=Split(input_Str,"~")

Comp_Char=inputBox("Enter The Comparing Character")

For int_i=0 to ubound(arrFruits)
	Redim Preserve Filter_arrFruits(index)
	if Mid(arrFruits(int_i),1,1)=Comp_Char then
		Filter_arrFruits(index)=arrFruits(int_i)
		index=index+1
	end if
next

if index=0 then
		Wscript.Echo "Their No Array Element Start With "&Comp_Char
Else
	Wscript.Echo " Array Element Start With "&Comp_Char&" = "&Join(Filter_arrFruits," , ")
End if