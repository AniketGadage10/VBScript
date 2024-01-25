OPTION EXPLICIT 

DIM A
A=10
DIM B
B=20

msgbox(" A >= 10 AND B >=10 = "&(A >= 10 AND B >=10))

msgbox(" A >= 10 AND B > 20 = "&(A >= 10 AND B > 20))


msgbox(" A >= 10 OR B >=10 = "&(A >= 10 OR B >=10))

msgbox(" A >= 10 OR B > 20 = "&(A >= 10 OR B > 20))



msgbox(" NOT (A >= 10 OR B >=10 ) = "&NOT(A >= 10 OR B >=10))

msgbox(" NOT (A >=10 AND  B > 20 ) = "&not(A >= 10 AND  B > 20))




msgbox(" A >= 10 XOR B >=10 = "&(A >= 10 XOR B >=10))

msgbox(" A >= 10 XOR B > 20 = "&(A >= 10 XOR B > 20))