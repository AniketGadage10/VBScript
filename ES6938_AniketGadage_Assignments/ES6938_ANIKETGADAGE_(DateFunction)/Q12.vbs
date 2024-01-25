'Give an example of FormatDateTime Function.

Option Explicit

Dim Sample_Date
Sample_Date=CDate("10/03/2000 06:30:00 AM")

MsgBox " Format No 0 :"&vbNewLine&"  "&FormatDateTime(Sample_Date,0)
MsgBox " Format No 1 :"&vbNewLine&"  "&FormatDateTime(Sample_Date)
MsgBox " Format No 2 :"&vbNewLine&"  "&FormatDateTime(Sample_Date,1)
MsgBox " Format No 3 :"&vbNewLine&"  "&FormatDateTime(Sample_Date,2)
MsgBox " Format No 4 :"&vbNewLine&"  "&FormatDateTime(Sample_Date,3)
MsgBox " Format No 5 :"&vbNewLine&"  "&FormatDateTime(Sample_Date,4)

								