Attribute VB_Name = "Module1"
Option Explicit

Public Password$
Public Head$

Sub Cipher(Txt$, Optional Rvalue, Optional A, Optional B)
Static R As Long
Static M As Long
Static N As Long
Const BigNum As Long = 32768
Dim i As Long, c As Long, d As Long
If IsMissing(Rvalue) = False Then
R = Rvalue
End If
If IsMissing(A) Then
If M = 0 Then M = 69
Else
M = (A * 4 + 1) Mod BigNum
End If
If IsMissing(B) Then
If N = 0 Then N = 47
Else
N = (B * 2 + 1) Mod BigNum
End If
For i = 1 To Len(Txt$)
c = Asc(Mid$(Txt$, i, 1))
Select Case c
Case 48 To 57
d = c - 48
Case 63 To 90
d = c - 53
Case 97 To 122
d = c - 59
Case Else
d = -1
End Select
If d >= 0 Then
R = (R * M + N) Mod BigNum
d = (R And 63) Xor d
Select Case d
Case 0 To 9
c = d + 48
Case 10 To 37
c = d + 53
Case 38 To 63
c = d + 59
End Select
Mid$(Txt$, i, 1) = Chr$(c)
End If
Next i
End Sub

Function Hash$(A$)
Dim i As Long, N As Long
Dim H$
For i = 1 To Len(A$)
N = N + Asc(Mid$(A$, i, 1))
N = (N * 1717 + 1717) Mod 1048576
Next i
For i = 1 To 7
N = (N * 997 + 997) Mod 1048576
Next i
H$ = Right$("0000" & Hex$(N), 4)
For i = 1 To Len(A$)
N = N + Asc(Mid$(A$, i, 1))
N = (N * 997 + 997) Mod 1048576
Next i
For i = 1 To 7
N = (N * 1717 + 1717) Mod 1048576
Next i
Hash$ = H$ & Right$("0000" & Hex$(N), 4)
End Function

