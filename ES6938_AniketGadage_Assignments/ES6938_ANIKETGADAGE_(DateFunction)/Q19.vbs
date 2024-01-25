'How can we display the week of the year and quarter of the year

Option Explicit

Dim Sample_Date
Sample_Date=CDate("25/04/2024 10:03:00 AM")

MsgBox " Week Of The Year : "&DatePart("ww",Sample_Date)
MsgBox " Quarter Of The Year : "&DatePart("q",Sample_Date)