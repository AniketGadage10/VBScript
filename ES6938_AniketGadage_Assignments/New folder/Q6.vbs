'Find Out the type of error in the code 


Dim arrCar(2)
	ON Error Resume Next
		arrCar(0)="Maruti"
		arrCar(1)="Tata"
		arrCar(5)="Mahindra"
		MsgBox "ERROR Number : "&Err.Number&" "&vbNewLine&" ERROR Description : "&Err.Description
		
		Err.Clear
		
	on Error Goto 0
	
	
'Output
'ERROR Number=9
'ERROR Description=Script Out Of Range 
