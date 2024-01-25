'Example's Of DatePart

Dim Sample_Date
'Sample_Date=CDate(Date&" "&Time)
Sample_Date=CDate("10/03/2000 10:03:00 AM")

msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition : DatePart('yyyy',Sample_Date)"&vbNewLine&" Output : "&DatePart("yyyy",Sample_Date))
msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition :  DatePart('m',Sample_Date)"&vbNewLine&" Output : "&DatePart("m",Sample_Date))
msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition :  DatePart('d',Sample_Date)"&vbNewLine&" Output : " &DatePart("d",Sample_Date))
msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition :  DatePart('y',Sample_Date)"&vbNewLine&" Output : " &DatePart("y",Sample_Date))
 msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition :  DatePart('w',Sample_Date)"&vbNewLine&" Output : " &DatePart("w",Sample_Date))
 msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition :  DatePart('ww',Sample_Date"&vbNewLine&" Output : " &DatePart("ww",Sample_Date))
 msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition :  DatePart('h',Sample_Date)"&vbNewLine&" Output : " &DatePart("h",Sample_Date))
 msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition : DatePart('n',Sample_Date)"&vbNewLine&" Output : " &DatePart("n",Sample_Date))
 msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition :  DatePart('s',Sample_Date)"&vbNewLine&" Output : " &DatePart("s",Sample_Date))
