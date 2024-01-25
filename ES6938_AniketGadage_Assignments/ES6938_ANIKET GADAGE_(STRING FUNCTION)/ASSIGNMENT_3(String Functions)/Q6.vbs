
var="ExpleoIndia"

Dim str:str=""

For i=1 to len(var) 
	s=Mid(var,i,1)
	x=cstr(count(s,var))
	str=str+s+" = "+x+" "&vbnewline

Next

Wscript.echo " Character That Constitute In String : "&var&"  = "&vbnewline&""&str


Function count(ch,v)
	c=0
	for j=1 to len(v)
		IF ch=MID(v,j,1) then
			c=c+1
		END IF
	next
	count = c
End Function














'var="ExpleoIndia"

'Dim str:str=""

'For i=1 to len(var) 
'	s=Mid(var,i,1)
'	if InStr(str,s)=0 then
'		str=str+s+" , "
'	end if
'Next

'Wscript.echo " Character That Constitute In String : "&var&" = "&str