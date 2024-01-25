

Dim fruit(3)

fruit(0)="Apple"
fruit(1)="Mango"
fruit(2)="Orange"
fruit(3)="GUVA"


for i=0 to ubound(fruit)
	msgbox(fruit(i))
next

for each f in fruit
	IF F="Orange" THEN
		eXIT FOR
	END IF
	msgbox(f)
next
