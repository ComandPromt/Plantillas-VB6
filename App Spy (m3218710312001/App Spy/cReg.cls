VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Function QueryValueEx(ByVal hKey As Long, ByVal szValueName As String, vValue As Variant) As Long
Dim Handle As Long, lrc As Long, lType As Long, lValue As Long, length As Long
Dim sValue As String
On Error GoTo ExError
    lrc = RegQueryValueExNULL(hKey, szValueName, 0&, lType, 0&, Handle)
    If lrc <> 0& Then Error 5
    Select Case lType
        Case (1):
            sValue = String(Handle, 0)
            lrc = RegQueryValueExString(hKey, szValueName, 0&, lType, sValue, Handle)
            If lrc = 0& Then vValue = sValue Else vValue = "Error"
        Case Else
            lrc = "ERROR"
    End Select
ExError:
    QueryValueEx = lrc
    Exit Function
End Function
Public Function qVal(hKey As Long, sKeyName As String, sValueName As String)
Dim lRetVal As Long, Handle As Long
Dim vValue As Variant
    lRetVal = RegOpenKeyEx(hKey, sKeyName, 0, &H20019, Handle)
    lRetVal = QueryValueEx(Handle, sValueName, vValue)
    qVal = vValue
    RegCloseKey (Handle)
End Function
Function GetKey(hKey As Long, Section As String, Keys() As String) As Long
Dim Handle As Long, lngResult As Long, lngCurIdx As Long
Dim lngValueLen As Long, lngData As Long, lngDataLen As Long
Dim strValue As String, strResult As String, Str1 As String
lngResult = RegOpenKeyEx(hKey, Section, 0&, &H20019, Handle)
If lngResult <> 0& Then Exit Function
lngCurIdx = 0
Do
    lngValueLen = 2000
    strValue = String(lngValueLen, 0)
    lngDataLen = 2000
    lngResult = RegEnumValue(Handle, lngCurIdx, ByVal strValue, lngValueLen, 0&, (1), ByVal lngData, lngDataLen)
    lngCurIdx = lngCurIdx + 1
    If lngResult = 0& Then
        strResult = Left(strValue, lngValueLen)
        Keys(lngCurIdx) = strResult
    Else
        Keys(lngCurIdx) = ""
    End If
Loop While lngResult = 0&
RegCloseKey (Handle)
GetKey = lngCurIdx - 1
End Function
Public Function DelKey(hKey As Long, sKeyName As String)
Dim lRetVal As Long
Dim Handle As Long
    lRetVal = RegOpenKeyEx(hKey, sKeyName, 0, &H3F, Handle)
    lRetVal = RegDeleteKey(hKey, sKeyName)
    RegCloseKey (Handle)
End Function
Public Function DelVal(hKey As Long, sKeyName As String, sValueName As String)
Dim Handle As Long
    RegOpenKeyEx hKey, sKeyName, 0, &H3F, Handle
    RegDeleteValue Handle, sValueName
    RegCloseKey (Handle)
End Function
