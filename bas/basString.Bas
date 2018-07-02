Attribute VB_Name = "basString"
Option Explicit

Private Declare Function IsCharAlpha Lib "user32" _
   Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
   
Private Declare Function IsCharAlphaNumeric Lib _
   "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long


'
' This function returns the Nth token in a string
'    Ex.  GetWord("This is a test.", " ", 2) = "is"
'
Public Function GetToken(s As String, token As String, ByVal Nth As Integer) As String
   Dim i As Integer
   Dim p As Integer
   Dim r As Integer

   If Nth < 1 Then
      GetToken = ""
      Exit Function
   End If

   r = 0

   For i = 1 To Nth
      p = r
      r = InStr(p + 1, s, token)
      If r = 0 Then
         If i = Nth Then
            GetToken = Mid$(s, p + 1, Len(s) - p)
         Else
            GetToken = ""
         End If
         Exit Function
      End If
   Next i

   GetToken = Mid$(s, p + 1, r - p - 1)
End Function
'
'  Returns an array to tokenized values
'  Ex:  GetTokens("This is a test.") = ({ "This", "is", "a", "test." })
'
Public Function GetTokens(sTxt As String, sToken As String) As Variant
    Dim iTokenLen As Integer
    Dim iTokenCnt As Integer
    Dim lOffset As Long
    Dim lPrevOffset As Long
    Dim aTokens() As String

    iTokenLen = Len(sToken)
    lOffset = InStr(sTxt, sToken)
    
    Do While lOffset > 0
        ReDim Preserve aTokens(iTokenCnt)
        If lOffset - lPrevOffset > 1 Then
            aTokens(iTokenCnt) = Mid$(sTxt, lPrevOffset + 1, lOffset - 1 - lPrevOffset)
        Else
            aTokens(iTokenCnt) = ""
        End If
        
        lPrevOffset = lOffset
        lOffset = InStr(lOffset + iTokenLen, sTxt, sToken)
        iTokenCnt = iTokenCnt + 1
    Loop
    
    ReDim Preserve aTokens(iTokenCnt)
    aTokens(iTokenCnt) = Mid$(sTxt, lPrevOffset + 1)
    GetTokens = CVar(aTokens)
End Function
' String functions.
' Converts a double to a string
' Note:  numbers after the decimal place
'        are ignored.
Function Int2String(ByVal l As Double) As String
   Dim tmp As String
   Dim str As String
   Dim i As Integer
   Dim j As Integer
   
   tmp = Format(l, "000000000000")
   str = ""
   
   ' Opps... it's more than 999 trillion
   ' One could easily add bigger number
   ' support.
   If Len(tmp) > 12 Then
      Int2String = ""
      Exit Function
   End If
   
   
   ' zero is a special case.
   ' you may want to change this to "no"
   ' as in "no dollars and 12/100" for writing
   ' checks.
   
   If Val(tmp) = 0 Then
      Int2String = "zero"
      Exit Function
   End If
   
   
   i = Val(Left$(tmp, 3))
   If i <> 0 Then
      GoSub do_hundreds
      str = str + " trillion"
   End If
   
   i = Val(Mid$(tmp, 4, 3))
   If i <> 0 Then
      GoSub do_hundreds
      str = str + " million"
   End If
   
   i = Val(Mid$(tmp, 7, 3))
   If i <> 0 Then
      GoSub do_hundreds
      str = str + " thousand"
   End If
   
      
   i = Val(Right$(tmp, 3))
   If i <> 0 Then
      GoSub do_hundreds
   End If
   
   Int2String = str
   Exit Function
   
   

do_hundreds:
   If i > 99 Then
      j = i
      i = i \ 100
      GoSub do_ones
      str = str + " hundred"
      i = j Mod 100
   End If

   If i <> 0 Then
      GoSub do_tens
   End If
   Return
   
do_tens:
   Select Case i Mod 100
      Case 90 To 99:
         str = str + " ninety"
         GoSub do_ones
      Case 80 To 89:
         str = str + " eighty"
         GoSub do_ones
      Case 70 To 79:
         str = str + " seventy"
         GoSub do_ones
      Case 60 To 69:
         str = str + " sixty"
         GoSub do_ones
      Case 50 To 59:
         str = str + " fifty"
         GoSub do_ones
      Case 40 To 49:
         str = str + " fourty"
         GoSub do_ones
      Case 30 To 39:
         str = str + " thirty"
         GoSub do_ones
      Case 20 To 29:
         str = str + " twenty"
         GoSub do_ones
         
      Case 19: str = str + " nineteen"
      Case 18: str = str + " eighteen"
      Case 17: str = str + " seventeen"
      Case 16: str = str + " sixteen"
      Case 15: str = str + " fifteen"
      Case 14: str = str + " fourteen"
      Case 13: str = str + " thirteen"
      Case 12: str = str + " twelve"
      Case 11: str = str + " eleven"
      Case 10: str = str + " ten"
      
      Case Else
         GoSub do_ones
   End Select
   Return
   
   
do_ones:
   If i < 10 Or i Mod 10 = 0 Then
      str = str + " "
   Else
      str = str + "-"
   End If
   
   Select Case i Mod 10
      Case 9: str = str + "nine"
      Case 8: str = str + "eight"
      Case 7: str = str + "seven"
      Case 6: str = str + "six"
      Case 5: str = str + "five"
      Case 4: str = str + "four"
      Case 3: str = str + "three"
      Case 2: str = str + "two"
      Case 1: str = str + "one"
   End Select
   
   Return
End Function


'
' Returns 0 if the string is alpha.
' otherwise returns the position of the first character
' that failed the test.
'
Public Function IsStringAlpha(s As String) As Long
   Dim i As Long
   
   For i = 1 To Len(s)
      If IsCharAlpha(Asc(Mid$(s, i, 1))) = 0 Then
         IsStringAlpha = i
         Exit Function
      End If
   Next i
   
   IsStringAlpha = 0
End Function

'
' Returns 0 if the string is alphaNumeric
' otherwise returns the position of the first character
' that failed the test.
'
Public Function IsStringAlphaNumeric(s As String) As Long
   Dim i As Long
   
   For i = 1 To Len(s)
      If IsCharAlphaNumeric(Asc(Mid$(s, i, 1))) = 0 Then
         IsStringAlphaNumeric = i
         Exit Function
      End If
   Next i
   
   IsStringAlphaNumeric = 0
End Function
'
' Returns 0 if the string is Numeric
' otherwise returns the position of the first character
' that failed the test.
'
Public Function IsStringNumeric(s As String) As Long
   Dim i As Long
   Dim j As Byte
   
   For i = 1 To Len(s)
      j = Asc(Mid$(s, i, 1))
      If IsCharAlphaNumeric(j) = 1 Then
         If IsCharAlpha(j) = 1 Then
            IsStringNumeric = i
            Exit Function
         End If
      Else
         IsStringNumeric = i
         Exit Function
      End If
   Next i
   
   IsStringNumeric = 0
End Function
'trim a string returned from a system function.
'ie. kill the 0.
Public Function STrim(s As String) As String
   Dim i As Integer
   Dim s2 As String
   
   s2 = Trim(s)
   i = InStr(s2, Chr$(0))
   
   If i > 0 Then
      s2 = Left$(s2, i - 1)
   End If
   
   STrim = s2
End Function


