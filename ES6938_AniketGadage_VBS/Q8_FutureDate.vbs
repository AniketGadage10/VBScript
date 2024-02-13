' 8. Write a program to get date which will come after 13 months from today

Option Explicit
 
	Dim Cur_date

	Cur_date=date
	
	MsgBox Cur_date&" AFTER 13 MONTHS "&DateAdd("m",13,Cur_date)
	