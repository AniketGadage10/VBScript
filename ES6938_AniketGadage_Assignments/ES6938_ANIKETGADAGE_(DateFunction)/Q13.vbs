'Give an example of DateSerial.

Option Explicit

Dim Sample_Day,Sample_Month,Sample_Year,Sample_Date
Dim Flag:Flag=True

ON Error Resume Next

While Flag
	Sample_Day=INT(InputBox("Enter Day"))
	Sample_Month=INT(InputBox("Enter Month"))
	Sample_Year=INT(InputBox("Enter Year"))
		If Err.Number=0 Then
			Sample_Date=DateSerial(Sample_Year,Sample_Month,Sample_Day)'This Function return the date in mm/dd/yyyy this format
			MsgBox Sample_Date
			Flag=False
		Else
			MsgBox "ERROR Number : "&Err.Number&"ERROR : "&Err.Description&vbNewLine&" Given Values Must Be In Integer"
			Flag=True
		End If
		Err.Clear
Wend

on Error Goto 0