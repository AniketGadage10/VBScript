'Display Today's Date,Month,Year

Option Explicit

Dim Today_DayName,Today_Month,Today_Year,Today_Day
Today_DayName=WeekdayName(Weekday(Date))
Today_Day=Day(Date)
Today_Month=Month(Date)
Today_Year=Year(Date)

MsgBox " Todays"&vbNewLine&" Day : "&Today_Day&" DayName : "&Today_DayName&" Month : "&Today_Month&" Year : "&Today_Year