' 2. Create 2 arrays and copy all values of both the into 3rd Array

Option Explicit

Dim IArray_1(),IArray_1ay_2(),ICombine_Array()
Dim Size_1,Size_2,Int_I,Index:Index=0
Int_I=0
Size_1=Int(InputBox("Enter The Size Of 1st IArray_1"))
Redim IArray_1(Size_1)

For Int_I=0 to Size_1
	IArray_1(Int_I)=Int(InputBox("Enter The "&Int_I&"TH Element Of IArray_1"))
next

Size_2=Int(InputBox("Enter The Size Of 2nd IArray_2"))
Redim IArray_1ay_2(Size_2)
For Int_I=0 to Size_2
	IArray_1ay_2(Int_I)=Int(InputBox("Enter The "&Int_I&"TH Element Of IArray_2"))
next

Redim ICombine_Array(ubound(IArray_1)+ubound(IArray_1ay_2)+1)


For Int_I = 0 to ubound(IArray_1)
	ICombine_Array(Index)=IArray_1(Int_I)
	Index=Index+1
next

For Int_I= 0 to ubound(IArray_1ay_2)
	ICombine_Array(Index)=IArray_1ay_2(Int_I)
	Index=Index+1
next


Wscript.ECHO " FIRST IArray_1AY "&JOIN(IArray_1," , ")&VbNewLine

Wscript.ECHO " SEIndexOND IArray_1AY "&JOIN(IArray_1ay_2," , ")&VbNewLine

Wscript.ECHO " NEW  IArray_1AY "&JOIN(ICombine_Array," , ")&VbNewLine