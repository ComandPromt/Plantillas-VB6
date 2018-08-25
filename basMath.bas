Attribute VB_Name = "basMath"
Option Explicit

'
'  Converts a number in any base from 2 to 36
'  to a long.
'
'  Note, this doesn't verify if the string
'  is a valid number in the given base.
'
Public Function Base2Long(s As String, ByVal nB As Integer) As Long
   Dim s2 As String
   Dim i As Long
   Dim j As Long
   Dim X As Long
   Dim n As Boolean
   Dim s3 As String
   
   If Len(s) < 1 Then
      Base2Long = 0
      Exit Function
   End If
   
   s2 = UCase(s)
   
   If Left$(s2, 1) = "-" Then
      n = True
      s2 = Right$(s2, Len(s2) - 1)
   Else
      n = False
   End If
   
   j = 1
   X = 0
   
   For i = Len(s2) To 1 Step -1
      s3 = Mid$(s2, i, 1)
      Select Case s3
      Case "0" To "9":
         X = X + j * (Asc(s3) - 48)
      Case "A" To "Z":
         X = X + j * (Asc(s3) - 55)
      End Select
      
      j = j * nB
   Next i
   
   If n Then
      X = -X
   End If
   
   Base2Long = X
End Function

'
'  Converts the number n to any base between 2 and 36
'
Public Function Long2Base(ByVal n As Long, ByVal nB As Integer) As String
  Dim s As String
  Dim nD As Integer
  Dim Negative As Boolean

  Negative = n < 0
  n = Abs(n)
  
  Do
    nD = n Mod nB
    If nD > 9 Then
       nD = nD + 7
    End If
    
    s = Chr$(48 + nD) & s
    n = n \ nB
  Loop Until n = 0
  
  If Negative Then
    s = "-" & s
  End If
  
  Long2Base = s
End Function


'
' Returns true if the number is a prime number.
' false if it is not.
'
' This should work reasonably well for small
' numbers (32-bits or less).  For larger numbers
' the Rabin-Miller test should be used.
'
Public Function IsPrime(ByVal n As Long) As Boolean
    Dim i As Long

    IsPrime = False
    
    If n <> 2 And (n And 1) = 0 Then Exit Function 'test if div 2
    If n <> 3 And n Mod 3 = 0 Then Exit Function 'test if div 3
    For i = 6 To Sqr(n) Step 6
        If n Mod (i - 1) = 0 Then Exit Function
        If n Mod (i + 1) = 0 Then Exit Function
    Next
    
    IsPrime = True
End Function
