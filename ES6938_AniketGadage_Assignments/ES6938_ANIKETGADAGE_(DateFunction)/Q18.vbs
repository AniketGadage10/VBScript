'How can we calculate the hour difference?

Option Explicit

Dim Sample_Start,Sample_End,Hour_Diff
Sample_Start=Cdate("01-Jan-2024 2:45:00 AM")
Sample_End=Cdate("01-Jan-2024 12:0:00 AM")
Hour_Diff=DateDiff("h",Sample_End,Sample_Start)
MsgBox " Hour Difference Between : "&Sample_End&" AND "&Sample_End&" = "&Hour_Diff