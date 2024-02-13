' 4. Write a Program to print below pattern
' 1
' 2 2
' 3 3 3Se
' 4 4 4 4
' 5 5 5 5 5

Option Explicit

	Dim iRow,Int_i
	Dim sStr

	iRow=5
	sStr=""
	
	For Int_i=1 to iRow
		sStr=sStr+String(Int_i,CStr(Int_i))&vbNewLine
	Next
	
	Wscript.Echo sStr