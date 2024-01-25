'Find Out the type of error in the code 

Dim arrCar(2)
arrCar(0)="Maruti"
arrCar(1)="Tata"
arrCar(5)="Mahindra"



'Output 
'Runtime Error
'Error : SubScript Out Of Range [Number:5]
'As our array  is fixed size and upper bound is 2 but we are assigning value at  Index 5 so it is giving out of range error
'Line : 6