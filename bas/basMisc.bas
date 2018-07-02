Attribute VB_Name = "basMisc"
Option Explicit
' Misc. functions.

'
'  Sets a given bit in num
'
Public Function SetBit(Num As Long, ByVal bit As Long) As Long
   If bit = 31 Then
      Num = &H80000000 Or Num
   Else
      Num = (2 ^ bit) Or Num
   End If
   
   SetBit = Num
End Function

'
'  clears a given bit in num
'
Public Function ClearBit(Num As Long, ByVal bit As Long) As Long
   If bit = 31 Then
      Num = &H7FFFFFFF And Num
   Else
      Num = ((2 ^ bit) Xor &HFFFFFFFF) And Num
   End If
   
   ClearBit = Num
End Function

'
'  Test if bit 0 to bit 31 is set.
'
Public Function IsBitSet(ByVal Num As Long, ByVal bit As Long) As Boolean
   IsBitSet = False
   
   If bit = 31 Then
      If Num And &H80000000 Then
         IsBitSet = True
      End If
   Else
      If Num And (2 ^ bit) Then
         IsBitSet = True
      End If
   End If
End Function

' Centers a form relative to the screen or
' another form
Public Sub CenterForm(f As Form, Optional f2 As Variant)
   If IsMissing(f2) Then
      f.Move (Screen.Width - f.Width) / 2, _
             (Screen.Height - f.Height) / 2
   Else
      ' If f is an MDI child in a MDI parent then
      ' center f within the parent.
      If f.MDIChild And Not f2.MDIChild Then
            f.Move ((f2.Width - f.Width) / 2), _
            ((f2.Height - f.Height) / 2)
      Else
         f.Move ((f2.Width - f.Width) / 2) + f2.Left, _
                ((f2.Height - f.Height) / 2) + f2.Top
      End If
   End If
End Sub







Public Function IsEven(ByVal i As Long) As Boolean
   IsEven = Not -(i And 1)
End Function






'
'  Returns true if the year is a leap year.
'     yr is either a date or an integer
'
Public Function IsLeapYear(yr As Variant) As Boolean
   If VarType(yr) = vbDate Then
      IsLeapYear = (Day(DateSerial(Year(yr), 2, 29)) = 29)
   Else
      IsLeapYear = (Day(DateSerial(yr, 2, 29)) = 29)
   End If
End Function

Public Function IsOdd(ByVal i As Long) As Boolean
   IsOdd = -(i And 1)
End Function

'
'  Scrambles the order of elements in an array.
'
Public Sub ShuffleArray(ByRef vArray As Variant, Optional startIndex As Variant, Optional endIndex As Variant)
    Dim i As Long
    Dim rndIndex As Long
    Dim Temp As Variant
    
    If IsMissing(startIndex) Then
       startIndex = LBound(vArray)
    End If
    
    If IsMissing(endIndex) Then
       endIndex = UBound(vArray)
    End If

    For i = startIndex To endIndex
        rndIndex = Int((endIndex - startIndex + 1) * Rnd() + startIndex)

        Temp = vArray(i)
        vArray(i) = vArray(rndIndex)
        vArray(rndIndex) = Temp
    Next i
End Sub
