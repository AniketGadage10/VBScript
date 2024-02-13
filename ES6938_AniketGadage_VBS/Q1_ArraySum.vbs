
'1. Write a program for addition of all elements of that array
'arrVal = array(5,3,5,6,1)
'O/P = 20

Option Explicit

	Dim iArr
	Dim iSum,Int_i

	iArr=Array(5,3,5,6,1)

		For Int_i=0 to UBound(iArr)
			iSum=iSum+iArr(Int_i)
		Next
		
		
	Wscript.Echo " Sum Of Arrray Element : "&iSum