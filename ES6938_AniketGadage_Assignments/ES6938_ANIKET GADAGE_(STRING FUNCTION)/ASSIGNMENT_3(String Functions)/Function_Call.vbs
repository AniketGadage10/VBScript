

DIM E:E=True

while E 
WSCRIPT.ECHO "1 : FIND NO OF VOWEL IN STRING "&VBNEWLINE&"2 : FIND UPPERCASE AND LOWEERCASE LETTER "&VBNEWLINE&"3 : SWAPPING OF TWO STRING "&VBNEWLINE&"4 :STRING REVERSE WITHOUT REVERSE FUNCTION "&VBNEWLINE&"5 : IF COUNT IN STRING "&VBNEWLINE&"6 : ALL CHARACTER COUNT IN STRING "&VBNEWLINE&"7 : REVERSE THE ORDER OF WORDS "&VBNEWLINE&"8 : FIND LENGTH WITHOUT USING LENGTH FUNCTION "&VBNEWLINE&"9 : EXIT "
var=int(InputBox("Enter the  Operation"))

SELECT CASE var
	
CASE 1
	str=inputbox("Enter the String")
	MsgBox(vowel_finder(str))
CASE 2
	str=InputBox("Enter the String")
	MsgBox(upper_lower_finder(str))
CASE 3
	str1=InputBox("Enter the 1st String")
	str2=InputBox("Enter the 2nd String")
	A=Swap_String(str1,str2)
CASE 4
	str=InputBox("Enter the String")
	MsgBox(ReverseStr(str))
CASE 5
	str=inputbox(" Enter The String With 'if'")
	MsgBox(If_Count(str))
CASE 6
	str=InputBox("Enter the String")
	MsgBox(Char_Count(str))
CASE 7
	str=InputBox("Enter the String")
	REVERSE_OdR(str)
CASE 8
	str=InputBox("Enter the String")
	MsgBox(STR_LENGTH(STR))
CASE 9
	E=False
CASE else
	MsgBox "Wrong Input Operation Value"

END SELECT

WEND



'FUNCTIONS

Function vowel_finder(var)
	
	Dim c:c=0

	For i=1 to len(var)
		s=Lcase(MID(var,i,1))

		if s="a" OR s="i" OR s="e" OR s="o" OR s="u" then
			c=c+1
		end if

	next

	vowel_finder=" Total Vowel In "&var&" = "&c

END Function



Function upper_lower_finder(var)
	Dim lower1,upper1,number1
	lower1=""
	upper1=""

		for i=1 to len(var)
			s=MID(var,i,1)

			if Asc(s)>=65 AND Asc(s)<=90 then
	 			lower1=lower1+s+" , "
			elseif 	 Asc(s)>=97 AND Asc(s)<=122 then
				upper1=upper1+s+" , "
			elseif  Asc(s)>=48 AND Asc(s)<=57 then
				number1=number1+s+" , "
			end if
		next
		
	upper_lower_finder=" Lower Case Letter : "&lower1&vbnewline&" Upper Case Letter : "&upper1&vbnewline&" Numbers : "&number1
End Function


Function Swap_String(str1,str2)

		temp=str1

		Wscript.echo " Before Swap : String1 = "&str1&" String2 = "&str2

		str1=Replace(str1,str1,str2)

		str2=Replace(str2,str2,temp)

		Wscript.echo " After Swap : String1 = "&str1&" String2 = "&str2
	Swap_String=""
End Function

Function ReverseStr(var)
	rev=""

	for i=len(var) to 1 step -1
		rev=rev+MID(var,i,1)
	next

 ReverseStr=" Before Reverse String : "&var&vbnewline& " After Reverse String : "&rev
End Function

Function If_Count(var)
	
	Dim c:c=0

	for i=1 to len(var)-1
		s=MID(var,i,2)
		if s="if" then
		c=c+1
		end if
	next

	If_Count=" Total If In String "&var&" = "&c 
End Function


Function count(ch,v)
	c=0
	for j=1 to len(v)
		IF ch=MID(v,j,1) then
			c=c+1
		END IF
	next
	count = c
End Function

Function Char_Count(var)
	Dim str:str=""

		For i=1 to len(var) 
			s=Mid(var,i,1)
			x=cstr(count(s,var))
			str=str+s+" = "+x+" "&vbnewline

		Next

	Char_Count=" Character That Constitute In String : "&var&"  = "&vbnewline&""&str
End Function




Function REVERSE_OdR(var)
	DIM i,j
	str=""

	for i=len(var) to 1 step -1
		n=InstrRev(var," ",i)
		s=Mid(var,n+1,len(var)-n-len(str))
		str=str+s+" "
		i=n-1
	next
	MsgBox(" Reverse Order Of String = "&str)
End Function




Function STR_LENGTH(var)

		dim s:s=""

		Dim count:count=0

		for i=1 to 100 
			c=Mid(str,i,1)
			if c="" then
			Exit For
			end if
		count=count+1
	next

	STR_LENGTH=" Count Of String : "&str&" = "&count
End Function

