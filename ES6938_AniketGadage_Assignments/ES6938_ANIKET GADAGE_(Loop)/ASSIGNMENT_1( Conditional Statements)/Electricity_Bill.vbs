
unitreading=Cint(inputbox(" Enter The Electricity Unit "))

If unitreading>=250 then

	bill=(unitreading*1.50)+((unitreading*1.50*20)/100)

ELSEIF unitreading<250 AND unitreading>100 then

	bill=(unitreading*1.20)+((unitreading*1.20*20)/100)

ELSEIF unitreading<=100 AND unitreading > 50 THEN

	bill=(unitreading*0.75)+((unitreading*0.75*20)/100)
	
ELSE 
	bill=(unitreading*0.50)+((unitreading*0.50*20)/100)
	
END IF

MsgBox(" Electricity Bill :- "&bill)

