
'var="ExPLEoInDia"


var=InputBox("Enter the String")
dim lower1,upper1,number1

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

Wscript.Echo " Lower Case Letter : "&lower1&vbnewline&" Upper Case Letter : "&upper1&vbnewline&" Numbers : "&number1