Option Explicit

' erweitert Excel um eine Arbeitstage-Formel mit ewigem Kalender der freien Tage (nur NRW)

Function ARBEITSTAGEINTEGRIERT(Datum As Date, Arbeitstage As Integer)
Dim Jahr As Integer
Dim Monat As Integer
Dim Tag As Integer
Dim Ostersonntag As Long
Dim Feiertage(12) As Long

Jahr = Year(Datum)
Monat = Month(Datum)
Tag = Day(Datum)

Feiertage(0) = DateSerial(Jahr, 1, 1)   'Neujahr

Ostersonntag = Eastersunday(CLng(Jahr)) 'Ostersonntag
Feiertage(1) = Ostersonntag
Feiertage(2) = Ostersonntag - 2         'Karfreitag
Feiertage(3) = Ostersonntag + 1         'Ostermontag

Feiertage(4) = DateSerial(Jahr, 5, 1)   'ErsterMai

Feiertage(5) = Ostersonntag + 39        'Christihimmelfahrt

Feiertage(6) = Ostersonntag + 49        'Pfingstsonntag
Feiertage(7) = Pfingstsonntag + 1       'Pfingstmontag

Feiertage(8) = Ostersonntag + 60        'Frohnleichnahm

Feiertage(9) = DateSerial(Jahr, 10, 3)  'Einheit

Feiertage(10) = DateSerial(Jahr, 11, 1) 'Allerheiligen

Feiertage(11) = DateSerial(Jahr, 12, 25) 'Weihnachten1
Feiertage(12) = DateSerial(Jahr, 12, 26) 'Weihnachten2

ULTIMO = Application.WorksheetFunction.WorkDay(Datum, Arbeitstage, Feiertage)

End Function

Public Function Eastersunday(Jahr As Long) As Date
Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long
  
  a = Jahr Mod 19
  b = Jahr \ 100
  c = (8 * b + 13) \ 25 - 2
  d = b - (Jahr \ 400) - 2
  e = (19 * (Jahr Mod 19) + ((15 - c + d) Mod 30)) Mod 30
  If e = 28 Then
    If a > 10 Then
      e = 27
    End If
  ElseIf e = 29 Then
    e = 28
  End If
  f = (d + 6 * e + 2 * (Jahr Mod 4) + 4 * (Jahr Mod 7) + 6) Mod 7
Eastersunday = DateSerial(Jahr, 3, e + f + 22)
  
End Function
