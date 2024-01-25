vno=int(inputbox("Enter The Number"))
sno=vno
sum=0
while vno<>0
	a= vno Mod 10
	vno=int(vno /10 )
	sum=sum+a
wend
Wscript.echo " SUM OF DIGIT IN "&sno&" IS "&sum