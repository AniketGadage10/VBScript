var1=inputbox("Enter the 3 digit no")

a= var1 MOD 10

var1=cint(cint(var1-a)/10)


b= var1 MOD 10
var1=cint(cint(var1-b)/10)



c= var1 MOD 10
var1=cint(cint(var1)/10)

d=(a*100)+(b*10)+c

Msgbox("Reverse no = "&d)
