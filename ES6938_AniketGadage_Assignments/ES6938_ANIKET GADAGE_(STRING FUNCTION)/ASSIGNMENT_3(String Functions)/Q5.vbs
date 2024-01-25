
'var="if + if =2if"

var=inputbox(" Enter The String ")

Dim c:c=0

for i=1 to len(var)-1
	s=MID(var,i,2)
	if s="if" then
	c=c+1
	end if
next

Wscript.Echo " Total If In String "&var&" = "&c 