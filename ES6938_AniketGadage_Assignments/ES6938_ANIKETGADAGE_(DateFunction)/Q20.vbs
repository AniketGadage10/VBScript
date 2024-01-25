'Display all Month and quarter of the year from the Given Date

Option Explicit

Dim Sample_Date,Int_I
Sample_Date=Date

Int_I=Month(Sample_Date)

For Int_I=Int_I to 12
	MsgBox " Month Name : "&MonthName(Month(Sample_Date),True)&" Quarter Of Year : "&DatePart("q",Sample_Date)
	Sample_Date=DateAdd("m",1,Sample_Date)
Next