' 5. write a program to check str1="UFT Traning" and str2="uft training" are
 ' equal or not, using strComp function
 
 
Option Explicit

	Dim sStr1,sStr2
	Dim iResult
	
	sStr1="UFT Traning"
	sStr2="uft training"
	
	iResult=StrComp(sStr1,sStr2)
	
	If iResult=0 Then
		MsgBox "Given Strings Are equal"	
	ElseIF iResult=-1 Then
		MsgBox sStr1&" Has Less Ascii Value Then "&sStr2
	Else
		MsgBox sStr1&" Has More Ascii Value Then "&sStr2
	END If