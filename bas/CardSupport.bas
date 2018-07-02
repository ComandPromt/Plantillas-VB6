Attribute VB_Name = "CardSupport"
Option Explicit

'This module has been included to avoid the problems associated
'with 16 and 32 bit versions of the Cards.dll.
'
'The strategy is to determine which of these DLL's is available
'on the target machine and to direct function calls to an
'appropriate wrapper function.
'This is not ideal and will cause some small performance hit

Private blnCard16Bit As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private function declarations                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Private Declare Function cdtInit32 Lib "cards.dll" Alias "cdtInit" ( _
    dX As Long, dY As Long) As Long
Private Declare Function cdtDrawExt Lib "cards.dll" (ByVal hDC As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, _
    ByVal ordCard As Long, ByVal iDraw As Long, ByVal clr As Long) As Long
Private Declare Function cdtDraw32 Lib "cards.dll" Alias "cdtDraw" (ByVal hDC As Long, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal iCard As Long, ByVal iDraw As Long, ByVal clr As Long) As Long
Private Declare Function cdtAnimate32 Lib "cards.dll" Alias "cdtAnimate" (ByVal hDC As Long, _
    ByVal iCardBack As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal iState As Long) As Long
Private Declare Function cdtTerm32 Lib "cards.dll" Alias "cdtTerm" () As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If you have the 16 Bit version of the Cards.dll (146KB for Windows ver 3.0)'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function cdtInit16 Lib "cards.dll" Alias "cdtInit" ( _
    dX As Integer, dY As Integer) As Integer
Private Declare Function cdtDraw16 Lib "cards.dll" Alias "cdtDraw" (ByVal hDC As Long, _
    ByVal X As Integer, ByVal Y As Integer, _
    ByVal iCard As Long, ByVal iDraw As Long, ByVal clr As Long) As Integer
Private Declare Function cdtTerm16 Lib "cards.dll" Alias "cdtTerm" () As Integer

''''''''''''''''''''''''''''''''''''''''
'  Wraper functions                    '
''''''''''''''''''''''''''''''''''''''''
Public Function cdtInit(vardX, vardY) As Variant
  If blnCard16Bit Then
     cdtInit = cdtInit16(vardX, vardY)
  Else
     cdtInit = cdtInit32(vardX, vardY)
  End If
End Function
Public Function cdtDraw(ByVal hDC As Long, ByVal X As Variant, ByVal Y As Variant _
          , lngCard As Long, lngdraw As Long, lngClr As Long) As Variant
          
  If blnCard16Bit Then
    cdtDraw = cdtDraw16(hDC, X, Y, lngCard, lngdraw, lngClr)
  Else
    cdtDraw = cdtDraw32(hDC, X, Y, lngCard, lngdraw, lngClr)
  End If
End Function
Public Function cdtTerm() As Variant
  If blnCard16Bit Then
    cdtTerm = cdtTerm16()
  Else
    cdtTerm = cdtTerm32()
  End If
End Function

'''''''''''''''''''''''''''''''''''''''''
'  Helper routines                      '
'''''''''''''''''''''''''''''''''''''''''

Public Function IsCards16() As Boolean
  IsCards16 = blnCard16Bit
End Function

Public Sub InitSupport()
  On Error GoTo er_InitSupport
  Call DetermineCards16Bit
  Exit Sub
er_InitSupport:
  Select Case Err.Number
    Case 53: 'File not found
      'this suggests that the cards DLL is not on the users system
      Err.Description = "Cards.dll File not found"
      Err.Raise
    Case Else:
      Err.Raise
  End Select
End Sub

Private Sub DetermineCards16Bit()
'The 16 bit version has a file length of 148528 bytes
'The NT 32 bit version has a file length of 156432 bytes.
'Since it is more likely that the 32 bit version will change slightly
'over time with new 32 bit builds of the OS, we'll
'check against the length compared to the 16 bit file.
  Dim strPath As String
  strPath = Space(256) 'initialize string
  Dim lngRet As Long
  Dim lngFileSize As Long
  On Error GoTo er_det16
  lngRet = GetWindowsDirectory(strPath, Len(strPath) - 1)
  strPath = Left(strPath, lngRet)
  lngFileSize = FileLen(strPath & "\cards.dll")
  If lngFileSize = 148528 Then 'length of the 16 bit DLL
    blnCard16Bit = True
  Else
    blnCard16Bit = False
  End If
  
Exit Sub
er_det16:
    strPath = Space(256)
    lngRet = GetSystemDirectory(strPath, Len(strPath) - 1)
    strPath = Left(strPath, lngRet)
    'Must check size here in case DLL does not exist at all on system
    lngFileSize = FileLen(strPath & "cards.dll")
    Resume Next
End Sub
