VAR =INPUTBOX("Enter the no")

ans = CINT(VAR) mod 2 

If ans = 0 Then 
msgbox (VAR&" IS EVEN NUMBER ")
Else
msgbox(VAR&" IS NOT EVEN NUMBER ")
END IF