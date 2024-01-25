'Input : arrFruits = Array("Apple","Grapes","Banana","Guava","Blueberries")
'Write a program which gives one array containg fruits name starting with "B"
Option Explicit

Dim arrFruits
Dim Ch_Char
Dim Int_I,Index,Filter_FruitArr()
Index=0
arrFruits = Array("Apple","Grapes","Banana","Guava","Blueberries")

Ch_Char=InputBox("Enter The Comparing Character")

For INT_I=0 to ubound(arrFruits)
	Redim Preserve Filter_FruitArr(Index)
	if MID(arrFruits(INT_I),1,1)=Ch_Char then
		Filter_FruitArr(Index)=arrFruits(INT_I)
		Index=Index+1
	end if
Next

If Index=0 then
	Wscript.Echo "Their No Array Element Start With "&Ch_Char
Else
	Wscript.Echo " Array Element Start With "&Ch_Char&" = "&Join(Filter_FruitArr," , ")
End If
