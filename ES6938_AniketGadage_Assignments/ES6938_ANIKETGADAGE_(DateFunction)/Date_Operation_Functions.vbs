
Option Explicit

Dim Select_Opr
Dim Flag:Flag=True

While Flag

	Wscript.Echo " 1: Todays' Date "&vbNewLine&" 2: Todays Date And Time"&vbNewLine&" 3: Todays Time"&vbNewLine&" 0: Exit"
	Select_Opr=Int(InputBox(" ENTER THE OPERATION NUMBER WANTS TO PERFORM "))
	
		Select Case Select_Opr
		case 0
			Flag=False
		Case 1
			Dim Cur_Date
			Cur_Date=Date
			MsgBox " Todays Date : "&Cur_Date
		Case 2
			Dim Cur_Date1,Cur_Time1
			Cur_Date1=Date
			Cur_Time1=Time
			Wscript.Echo " Todays Date_Time : "&Cur_Date1&"  "&Cur_Time1
		Case 3
			Dim Cur_Time3
			Cur_Time3=Time
			MsgBox " Todays Time : "&Cur_Time3
		
		

		End Select

Wend