

a=INPUTBOX("Enter the First Subject Marks")
b=INPUTBOX("Enter the Second Subject Marks")
c=INPUTBOX("Enter the Third Subject Marks")
d=INPUTBOX("Enter the Fourth Subject Marks")
e=INPUTBOX("Enter the Fith Subject Marks")

total=CInt(a)+CInt(b)+CInt(c)+CInt(d)+CInt(e)

msgbox(total)

percentage=((total/500)*100)

msgbox(percentage)

if percentage >=90 then 
    msgbox(" percentage = "&percentage&" Grade = A")
ELSEIF percentage >=80 then 
    msgbox(" percentage = "&percentage&" Grade = B")
ELSEIF percentage >=70 then 
    msgbox(" percentage = "&percentage&" Grade = C")
ELSEIF percentage >=60 then 
    msgbox(" percentage = "&percentage&" Grade = D")
ELSEIF percentage >=50 then 
    msgbox(" percentage = "&percentage&" Grade = E")
ELSE
    msgbox(" percentage = "&percentage&" Grade = F")
end if