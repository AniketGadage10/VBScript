' 10. Write a program to read first 5 Rows data from excels 2nd sheet and print it 
' (consider data available in only 1st column)


Option Explicit

	Dim xl_obj,xlworkbook_obj,xlworksheet_obj

	Dim iROW,Int_i

	iROW=5

	Set xl_obj=CreateObject("Excel.Application")

	xl_obj.Visible=True

	xl_obj.Application.DisplayAlerts=False


	set xlworkbook_obj = xl_obj.Workbooks.open "C:\Users\agadage\Desktop\test.xlsx"


	set xlworksheet_obj=xlworkbook_obj.Worksheets(2)


			For Int_i=1 to iROW 
					MsgBox xlworksheet_obj.cells(Int_i,1).value
			Next

	xlworkbook_obj.Save

	xl_obj.Quit

	set xl_obj=Nothing
	Set xlworkbook_obj=Nothing
	Set xlworksheet_obj=Nothing





