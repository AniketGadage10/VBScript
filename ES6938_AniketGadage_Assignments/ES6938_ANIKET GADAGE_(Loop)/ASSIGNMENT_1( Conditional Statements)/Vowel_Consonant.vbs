s=INPUTBOX(" Enter The Alphabate")

IF Mid(s,1,1)="A" OR Mid(s,1,1)="a" THEN
	WScript.echo " Alphabate "&CH&" Is a Vowel "
else
	WScript.echo " Alphabate "&CH&" Is a Consonant "
End IF



'Mid(String,Start,end) return the String OR CHar
's="abcs"  Mid(s,1,1) return a