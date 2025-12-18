Sub hardo_AG()  ' for ActiGraph epoch data  VBA Excel

Dim ac(10) As Double    ' vector for results

' settings, default Excel sheet 2  range("..")

pat = Worksheets(2).Range("B1").Value   ' file path for epoch meter files
LPA = Worksheets(2).Range("B4").Value   ' > Light activity counts per minute, < sedentary time
MPA = Worksheets(2).Range("B5").Value   ' > moderate activity counts per minute
VPA = Worksheets(2).Range("B6").Value   ' > vigorous activity counts per minute
EXC = Worksheets(2).Range("B7").Value   ' spurious over the limit data, excluded

vm = Worksheets(2).Range("B9").Value    ' vector magnitude or 1D
svm = Worksheets(2).Range("B10").Value  ' vector magnitude sedentary data (below LPA)

NWT = Worksheets(2).Range("B12").Value  '  minimum limit for excluded day (e.g. 600 min/day)
EX0 = Worksheets(2).Range("B13").Value  '  minimum limit for continuous zeros (e.g. 30 min)
CST = Worksheets(2).Range("B14").Value  ' limit of continuous sedentary time (e.g. 10 min)

A1 = Worksheets(2).Range("B16").Value   ' nighttime, going to bed
A2 = Worksheets(2).Range("C16").Value   ' nighttime, waking up

e = 60   '  assumption epoch (s)

Set a = Worksheets(1).Range("A2")   ' set the starting point for writing (excel)

If a.Value <> "" Then  ' If there is previous data, let's start from where we left off.
  Set a = a.Offset(1, 0)
End If

If Mid(pat, Len(pat), 1) <> "\" Then pat = pat & "\"  ' If the slash is missing from the end of the file path, add it.

csv = Dir(pat & "*.csv", 0)

Do While csv <> ""

z0 = -1  '  previous date, set -1

f = 0  ' overnight filter
n = 0
cc = 0
cs = 0

For i = 0 To UBound(ac)
  ac(i) = 0
Next i

Open pat & csv For Input As #1

Do Until EOF(1)

Line Input #1, l1

x = Split(l1, ",")

If Val(x(0)) = 0 Then GoTo NXT

dx = Split(x(0), " ")

d = DateValue(dx(0))

t = 0
If UBound(dx) > 0 Then
  t = TimeValue(Replace(dx(1), ".", ":"))   '  if the time string contains "." then change it to ":"
End If

Z = d + t  ' moment in time

If t = A1 Then f = 1
If t = A2 Then f = 0

If z0 >= 0 Then e = Round((Z - z0) * 24 * 60 * 60, 0)   ' determine the correct epoch using consecutive dates

If vm Then
  c = Sqr(Val(x(1)) ^ 2 + Val(x(2)) ^ 2 + Val(x(3)) ^ 2)
Else
  c = Val(x(1))
End If

If svm = False Then
  c0 = Val(x(1))
Else
  c0 = c
End If



s = Val(x(4))  ' steps

If Int(Z) <> Int(z0) And z0 > -1 Then  ' the moment when the day changes
  a.Value = csv
  a.Offset(0, 1).Value = Int(z0)
  
  ac(1) = ac(0) - ac(2) - ac(3) - ac(4)
  
  For i = 0 To UBound(ac)
    a.Offset(0, 2 + i).Value = ac(i)
    ac(i) = 0
  Next i
  
  
  Set a = a.Offset(1, 0)
  
  If a.Offset(-1, 2).Value < NWT Then
    a.Offset(-1, 0).EntireRow.Delete
  End If
  
End If

If f = 0 Then
  If c = 0 Then
    n = n + e / 60
  Else
    If n < EX0 Then
      ac(0) = ac(0) + (n + e / 60)
    Else
      If cs - n >= CST Then  ' continuous sedentary time when ex0 has been reached
        ac(8) = ac(8) + 1
        ac(9) = ac(9) + cs - n
      End If
      cs = 0  ' continuous sedentary activity is reset to zero
    End If
    n = 0
  End If
End If


If c > EXC * e / 60 Then   ' calculate different activity levels, possibly also at night
  ac(5) = ac(5) + 1 * e / 60
  ac(0) = ac(0) + 1 * e / 60
ElseIf c > VPA * e / 60 Then
  ac(4) = ac(4) + 1 * e / 60
ElseIf c > MPA * e / 60 Then
  ac(3) = ac(3) + 1 * e / 60
ElseIf c0 > LPA * e / 60 Then
  ac(2) = ac(2) + 1 * e / 60
Else
  If f = 1 Then ac(10) = ac(10) + 1 * e / 60  ' calculate the time of night
End If

ac(6) = ac(6) + s
ac(7) = ac(7) + c

If f = 0 Then

  cz = cz + c0
  cc = cc + 1

  If cc Mod 60 / e = 0 Then  ' calculate continuous sedentary time, default 60 s segments

  If cz <= LPA Then
    cs = cs + 1
  Else
    If cs > CST Then
      ac(8) = ac(8) + 1
      ac(9) = ac(9) + cs
    End If
    cs = 0
  End If
  cz = 0
  End If
End If

z0 = Z

NXT:

Loop

Close #1
csv = Dir  ' next file

Loop

End Sub
