Attribute VB_Name = "Module1"
' En/De-Crypt a string...
' Coded By MAGiC MANiAC^mTo
' More?... http://home.kabelfoon.nl/~mto
'
Function StrCrypt(ByVal sStr As String, ByVal Key As Long, ByVal Encrypt As Boolean) As String
  Dim lTmp1 As Long
  Dim lTmp2 As Long
  Dim lTmp3 As Long
  Dim lTmp4 As Long
  Dim lTmp5 As Long
  Dim lTmp6 As Long
  Dim sTmp1 As String
  lTmp1 = Len(sStr)
  sTmp1 = Space(lTmp1)
  ReDim lTmp7(lTmp1) As Long
  lTmp3 = 11 + (Key Mod 233)
  lTmp4 = 7 + (Key Mod 239)
  lTmp5 = 5 + (Key Mod 241)
  lTmp6 = 3 + (Key Mod 251)
  For lTmp2 = 1 To lTmp1: lTmp7(lTmp2) = Asc(Mid(sStr, lTmp2, 1)): Next
  If Encrypt Then
    For lTmp2 = 2 To lTmp1
      lTmp7(lTmp2) = lTmp7(lTmp2) Xor lTmp7(lTmp2 - 1) Xor ((lTmp3 * lTmp7(lTmp2 - 1)) Mod 256)
    Next
    For lTmp2 = lTmp1 - 1 To 1 Step -1
      lTmp7(lTmp2) = lTmp7(lTmp2) Xor lTmp7(lTmp2 + 1) Xor (lTmp4 * lTmp7(lTmp2 + 1)) Mod 256
    Next
    For lTmp2 = 3 To lTmp1
      lTmp7(lTmp2) = lTmp7(lTmp2) Xor lTmp7(lTmp2 - 2) Xor (lTmp5 * lTmp7(lTmp2 - 1)) Mod 256
    Next
    For lTmp2 = lTmp1 - 2 To 1 Step -1
      lTmp7(lTmp2) = lTmp7(lTmp2) Xor lTmp7(lTmp2 + 2) Xor (lTmp6 * lTmp7(lTmp2 + 1)) Mod 256
    Next
    For lTmp2 = 1 To lTmp1
      Mid(sTmp1, lTmp2, 1) = Chr(lTmp7(lTmp2))
    Next
  Else
    For lTmp2 = 1 To lTmp1 - 2
      lTmp7(lTmp2) = lTmp7(lTmp2) Xor lTmp7(lTmp2 + 2) Xor (lTmp6 * lTmp7(lTmp2 + 1)) Mod 256
    Next
    For lTmp2 = lTmp1 To 3 Step -1
      lTmp7(lTmp2) = lTmp7(lTmp2) Xor lTmp7(lTmp2 - 2) Xor (lTmp5 * lTmp7(lTmp2 - 1)) Mod 256
    Next
    For lTmp2 = 1 To lTmp1 - 1
      lTmp7(lTmp2) = lTmp7(lTmp2) Xor lTmp7(lTmp2 + 1) Xor (lTmp4 * lTmp7(lTmp2 + 1)) Mod 256
    Next
    For lTmp2 = lTmp1 To 2 Step -1
      lTmp7(lTmp2) = lTmp7(lTmp2) Xor lTmp7(lTmp2 - 1) Xor (lTmp3 * lTmp7(lTmp2 - 1)) Mod 256
    Next
    For lTmp2 = 1 To lTmp1
      Mid(sTmp1, lTmp2, 1) = Chr(lTmp7(lTmp2))
    Next
  End If
  StrCrypt = sTmp1
End Function
