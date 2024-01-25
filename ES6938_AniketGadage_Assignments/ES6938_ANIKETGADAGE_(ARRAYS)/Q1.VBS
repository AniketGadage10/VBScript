'Write a program to sort an array of size 5 in decending order

option Explicit
Dim str
Dim iarr
Dim inti,intj,temp
str="45,20,80,90,40,70"
iarr=Split(str,",")

for inti=0 to ubound(iarr)
	for intj=inti+1 to ubound(iarr)
		if iarr(inti)<=iarr(intj) then
		temp=iarr(inti)
		iarr(inti)=iarr(intj)
		iarr(intj)=temp
		end if
	next
next

	WSCRIPT.ECHO Join(iarr,",")

