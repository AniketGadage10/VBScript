Option Explicit

Dim xl_obj,xlworkbook_obj,xlworksheet_obj

DIM i,j

set xl_obj=CreateObject("Excel.Application")

xl_obj.Visible=True

xl_obj.Application.DisplayAlerts=False

' xl_obj.Workbooks.add

' xl_obj.ActiveWorkBook.saveas "D:\test.xlsx"

set xlworkbook_obj=xl_obj.Workbooks.open("D:\test.xlsx")

' set xlworksheet_obj=xlworkbook_obj.Worksheets.add

' xlworksheet_obj.name="Master"

'Set xlworksheet_obj=xlworkbook_obj.Worksheets(1)

'set xlworksheet_obj=xlworkbook_obj.Worksheets(2)

set xlworksheet_obj=xlworkbook_obj.Sheets("Master")

xlworksheet_obj.cells(3,1).value="ID"
xlworksheet_obj.cells(3,2).value="NAME"
xlworksheet_obj.cells(3,3).value="LOCATION"

xlworksheet_obj.cells(4,1).value=1
xlworksheet_obj.cells(4,2).value="ANIKET"
xlworksheet_obj.cells(4,3).value="PUNE"

xlworkbook_obj.save

' msgbox xlworksheet_obj.usedrange.rows.Count

' msgbox xlworksheet_obj.usedrange.columns.Count

xlworksheet_obj.rows(2).Delete

xlworkbook_obj.Sheets("Sheet1").Delete()

xl_obj.quit

Set xl_obj=Nothing
Set xlworkbook_obj=Nothing
Set xlworksheet_obj=Nothing