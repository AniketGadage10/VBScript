
Option Explicit
Dim Age
Dim Flag:Flag=True
ON Error Resume Next
While Flag
	Age=INT(InputBox("Enter Your Age"))
		If Err.Number=0 Then
			MsgBox "Thanks For Registration"
			Flag=False
		Else
			MsgBox "ERROR Number : "&Err.Number&"ERROR : "&Err.Description&vbNewLine&" Age Must In Integer"
			Flag=True
		End If
		Err.Clear
Wend

on Error Goto 0

