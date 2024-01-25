
'var="Expleo India Infosystem"
var=InputBox(" Enter The String ")
DIM i,j
str=""

for i=len(var) to 1 step -1
	n=InstrRev(var," ",i)
	s=Mid(var,n+1,len(var)-n-len(str))
	str=str+s+" "
	i=n-1
next

Wscript.Echo " Reverse Order Of String = "&str














'var="Expleo India Infosystem"
'DIM i,j
'str=""
'i=1
'for i=1 to len(var)
'	s=""
'	for j=i to len(var)
'		c=MID(var,j,1)
'		if c=" " then
'			Exit For
'		end if
'		s=s+c		
'	next
'	msgbox(s&vbnewline)
'	i=j
'	str=str+Strreverse(s)+" "
'next

'Wscript.Echo str