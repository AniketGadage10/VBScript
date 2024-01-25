'str="ExploEx"

str=InputBox(" Enter The String")

dim s:s=""

Dim count:count=0

for i=1 to 100 
	c=Mid(str,i,1)
	if c="" then
		Exit For
	end if
	count=count+1
next

msgbox(" Count Of String : "&str&" = "&count)
