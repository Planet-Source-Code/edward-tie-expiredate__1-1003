<div align="center">

## Expiredate


</div>

### Description

Expiredate, it's very usefull for makers of shareware. It can worked for 30 days long or 60 days long. you can make new demo for. After expiredate your programma will not worked until customer pay to you.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Edward Tie](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/edward-tie.md)
**Level**          |Advanced
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/edward-tie-expiredate__1-1003/archive/master.zip)





### Source Code

```
Option Explicit
Dim day1 As Integer
Dim month1 As Integer
Dim basis As Long
Dim schrikbasis As Long
Dim e As Long
Dim year1 As Long
Dim moncode As Integer
Dim ff As Integer
Private Sub Form_Load()
' Expiredate(tm) 1.2 for freeware. It's usefull for makers of a kind of demo and shareware.
' Copyright(c) 1998-1999,
'
' Expire day, month, year , total day
' If you will make 30-day trial software then you can put total day
' Example: day1,month1,year1, 30
' Support is limited. See to www.tcsoftware.com
'
month1 = Month(Date)
year1 = Year(Date)
day1 = Day(Date)
Tdate$ = format(Date$, "DD/MM/YYYY")
Call expiredate(day1, month1, year1, 30)
If Mid(Tdate$, 7) > year1 Then GoTo diened
If Mid(Tdate$, 7) = year1 Then
 If Left(Mid(Tdate$, 4), 2) = month1 Then If Left(Tdate$, 2) > day1 Then GoTo
diened
 If Left(Mid(Tdate$, 4), 2) > month1 Then GoTo diened
 end if
goto er7
 diened:
 MsgBox "Old version of Syscal has been expired!"
er7:
Label1.Caption = Str(day1) + "-" + Str(month1) + "-" + Str(Year(Date))
End Sub
Sub expiredate(day1 As Integer, month1 As Integer, year1 As Long, expireday As Integer)
Dim moncode As Integer
Dim ff As Long
Dim basis As Long
Dim schrikbasis As Long
Dim e As Long
day1 = day1 + expireday
start:
moncode = 1
For ff = 1 To 7
 If month1 = moncode Then
 If day1 > 31 Then
 day1 = day1 - 31: month1 = month1 + 1
 If month1 = 13 Then
 year1 = year1 + 1: month1 = 1: GoTo eind
 Else: GoTo eind
 End If
 Else: Exit Sub
End If
End If
If moncode = 1 Then moncode = 3: GoTo st1
If moncode = 7 Then moncode = 8: GoTo st1
moncode = moncode + 2
st1:
Next ff
moncode = 4
ff = 0
For ff = 1 To 5
If month1 = moncode Then
 If day1 > 30 Then
 day1 = day1 - 30: month1 = month1 + 1: GoTo eind
 Else: Exit Sub
 End If
End If
If moncode = 6 Then moncode = 9: GoTo st2
moncode = moncode + 2
st2:
Next ff
basis = 1980
schrikbasis = 2000
For e = 1 To 32000
If year1 = schrikbasis Then GoTo gewoon
If basis = schrikbasis Then schrikbasis = schrikbasis + 400
If year1 = basis Then If Month(Date) = 2 Then If day1 > 29 Then day1 = day1 - 29: month1 = month1 + 1: GoTo eind
basis = basis + 4
Next e
gewoon:
If month1 = 2 Then
If day1 > 28 Then
 day1 = day1 - 28: month1 = month1 + 1
 End If
 Else: Exit Sub
End If
eind:
GoTo start
eind1:
End Sub
```

