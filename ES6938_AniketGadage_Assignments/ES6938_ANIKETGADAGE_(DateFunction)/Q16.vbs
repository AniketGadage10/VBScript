'How can we Display Hour,Minute,Second

Option Explicit

Dim Cur_Hour,Cur_Minute,Cur_Second

Cur_Hour=Hour(now())
Cur_Minute=Minute(now())
Cur_Second=Second(now())

MsgBox " Current Hour : "&Cur_Hour&"   Current Minute : "&Cur_Minute&"   Current Second : "&Cur_Second

MsgBox " Current Hour : "&DatePart("h",now())&"   Current Minute : "&DatePart("n",now())&"   Current Second : "&DatePart("s",now())