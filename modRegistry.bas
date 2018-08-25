Attribute VB_Name = "modRegistry"
Option Explicit

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Public Const REG_SZ = 1

Public Enum Hkeys
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum

Public Function SetDefaultValue(hKey As Hkeys, KeyPath As String, Data As String)

    Dim Path    As String
    
    If KeyPath = "" Then
        Path = vbNullString
    Else
        Path = KeyPath
    End If
    
    Call RegSetValue(hKey, KeyPath, REG_SZ, Data, Len(Data))
    
End Function

Public Function SetValue(hKey As Hkeys, KeyPath As String, ValueName As String, Data As String)

    Dim KeyID   As Long
    
    Call RegOpenKey(hKey, KeyPath, KeyID)
    Call RegSetValueEx(KeyID, ValueName, &O0, REG_SZ, ByVal Data, Len(Data))
    Call RegCloseKey(KeyID)
    
End Function

Public Function GetDefaultValue(hKey As Hkeys, KeyPath As String) As String

    Call RegQueryValue(hKey, KeyPath, GetDefaultValue, 255)
    
End Function

Public Function GetValue(hKey As Hkeys, KeyPath As String, ValueName As String) As String

    Dim KeyID   As Long
    
    Call RegOpenKey(hKey, KeyPath, KeyID)
    Call RegQueryValueEx(KeyID, ValueName, &O0, REG_SZ, GetValue, 255)
    Call RegCloseKey(KeyID)
    
End Function

Public Function DeleteValue(hKey As Hkeys, KeyPath As String, ValueName As String)

    Dim KeyID   As Long
    
    Call RegOpenKey(hKey, KeyPath, KeyID)
    Call RegDeleteValue(KeyID, ValueName)
    Call RegCloseKey(KeyID)
    
End Function

Public Function DeleteKey(hKey As Hkeys, KeyPath As String, KeyName As String)

    Dim KeyID   As Long
    
    Call RegOpenKey(hKey, KeyPath, KeyID)
    Call RegDeleteKey(KeyID, KeyName)
    Call RegCloseKey(KeyID)
    
End Function

Public Function CreateKey(hKey As Hkeys, KeyPath As String, KeyName As String) As Long

    Dim KeyID As Long
    
    Call RegOpenKey(hKey, KeyPath, KeyID)
    Call RegCreateKey(KeyID, KeyName, CreateKey)
    Call RegCloseKey(KeyID)
    
End Function
