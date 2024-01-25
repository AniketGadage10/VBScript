
'var="Expleo"
var=InputBox(" Enter The String ")
rev=""

for i=len(var) to 1 step -1
	rev=rev+MID(var,i,1)
next

Wscript.Echo " Before Reverse String : "&var&vbnewline& " After Reverse String : "&rev
