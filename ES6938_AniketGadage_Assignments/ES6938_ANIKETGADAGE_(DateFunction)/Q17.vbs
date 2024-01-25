'Difference between weekday and weekdayName?
'weekday : Function returns the value from 1 to 7
'Each value represent the week day form Sunday to Saturday
'Return Type is Integer
'weekDayName : Function returns the Name Of The Week using weekday value
'return type is String

Option Explicit

Dim Sample_Date
Sample_Date=CDate("25/01/2024 10:03:00 AM")

Select CASE weekday(Sample_Date)
	CASE 1
		MsgBox " Week Day : "&weekday(Sample_Date)&" Week Name : "&weekdayName(weekday(Sample_Date))
	CASE 2
		MsgBox " Week Day : "&weekday(Sample_Date)&" Week Name : "&weekdayName(weekday(Sample_Date))
	CASE 3
		MsgBox " Week Day : "&weekday(Sample_Date)&" Week Name : "&weekdayName(weekday(Sample_Date))
	CASE 4
		MsgBox " Week Day : "&weekday(Sample_Date)&" Week Name : "&weekdayName(weekday(Sample_Date))
	CASE 5
		MsgBox " Week Day : "&weekday(Sample_Date)&" Week Name : "&weekdayName(weekday(Sample_Date))
	CASE 6
		MsgBox " Week Day : "&weekday(Sample_Date)&" Week Name : "&weekdayName(weekday(Sample_Date))
	CASE 7
		MsgBox " Week Day : "&weekday(Sample_Date)&" Week Name : "&weekdayName(weekday(Sample_Date))
	CASE Else
		MsgBox " SYSTEM ERROR "
end Select