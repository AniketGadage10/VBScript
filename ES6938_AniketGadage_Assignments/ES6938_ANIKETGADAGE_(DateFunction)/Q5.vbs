'Example's Of DateAdd

Option Explicit

Dim Sample_Date
Sample_Date=CDate("10/03/2000 10:03:00 AM")

msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition : Dateadd('yyyy',1,Sample_Date)"&vbNewLine&" Output : "&Dateadd("yyyy",1,Sample_Date))
msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition :  Dateadd('m',1,Sample_Date)"&vbNewLine&" Output : "&Dateadd("m",1,Sample_Date))
msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition :  Dateadd('d',1,Sample_Date)"&vbNewLine&" Output : " &Dateadd("d",1,Sample_Date))
msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition :  Dateadd('y',1,Sample_Date)"&vbNewLine&" Output : " &Dateadd("y",1,Sample_Date))
 msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition :  Dateadd('w',2,Sample_Date)"&vbNewLine&" Output : " &Dateadd("w",2,Sample_Date))
 msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition :  Dateadd('ww',1,Sample_Date"&vbNewLine&" Output : " &Dateadd("ww",1,Sample_Date))
 msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition :  Dateadd('h',2,Sample_Date)"&vbNewLine&" Output : " &Dateadd("h",2,Sample_Date))
 msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition : Dateadd('n',3,Sample_Date)"&vbNewLine&" Output : " &Dateadd("n",3,Sample_Date))
 msgbox(" Input Date : "&Sample_Date&vbNewLine&" Condition :  Dateadd('s',4,Sample_Date)"&vbNewLine&" Output : " &Dateadd("s",4,Sample_Date))
