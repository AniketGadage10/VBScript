'3. Write a program which will take automation tool name as a input and show its 
'License cost as a output
'e.g UFT - 4 lac , Tosca - 2 lac , TestComplete - 5 lac

Option Explicit

	Dim sToolName
	MsgBox("Automation Tool"&vbNewLine&"1:UFT"&vbNewLine&"2:Tosca"&vbNewLine&"3:TestComplete")
	sToolName=UCase(InputBox("Enter The Automation Tool Name"))
	
	Select Case sToolName
		Case "UFT",1
				MsgBox "Automation Tool UFT With License cost 4 lac"
		Case "TOSCA",2
				MsgBox "Automation Tool TOSCA With License cost 2 lac"
		Case "TESTCOMPLETE",3
				MsgBox "Automation Tool TESTCOMPLETE With License cost 5 lac"
		Case Else
			MsgBox "Select The Proper Automation tool"
	end Select

