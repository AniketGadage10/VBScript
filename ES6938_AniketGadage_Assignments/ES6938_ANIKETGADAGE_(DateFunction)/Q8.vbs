'Use TimeSerial function returns the time for a specific Hour,Minute,Second

Option Explicit

Dim Sample_Hour,Sample_Minute,Sample_Second
Dim Flag:Flag=True

ON Error Resume Next

While Flag
	Sample_Hour=INT(InputBox("Enter Hours"))
	Sample_Minute=INT(InputBox("Enter Minute"))
	Sample_Second=INT(InputBox("Enter Second"))
		If Err.Number=0 Then
			MsgBox " Time : "&TimeSerial(Sample_Hour,Sample_Minute,Sample_Second)
			Flag=False
		Else
			MsgBox "ERROR Number : "&Err.Number&"ERROR : "&Err.Description&vbNewLine&" Given Values Must Be In Integer"
			Flag=True
		End If
		Err.Clear
Wend

on Error Goto 0