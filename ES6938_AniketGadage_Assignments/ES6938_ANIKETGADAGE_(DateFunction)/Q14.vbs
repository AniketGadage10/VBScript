'How Can we converts the given input string to a Date OR Time

Option Explicit

Dim Input_Date,Input_Time,Sample_Time,Sample_Date

	Input_Date=InputBox("Enter Day/Month/Year")
			Sample_Date=CDate(Input_Date)
			MsgBox " Date : "&Sample_Date
	Input_Time=InputBox("Enter HH:MM:SS AM/PM")
			Sample_Time=CDate(Input_Time) 
			MsgBox " Time : "&Sample_Time
