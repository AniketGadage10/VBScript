'var="ExpleoIndia"

var=inputbox("Enter the String")

Dim c:c=0

For i=1 to len(var)
	s=Lcase(MID(var,i,1))

	if s="a" OR s="i" OR s="e" OR s="o" OR s="u" then
		c=c+1
	end if

next

Wscript.echo " Total Vowel In "&var&" = "&c