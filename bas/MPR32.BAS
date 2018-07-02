Attribute VB_Name = "basMPR"
' Module Name               basMPR
' Module Description        Module which contains ALL Declarations
'                           for API calls in the basMPR Library
'
Option Explicit

Declare Function WNetAddConnection _
    Lib "mpr.dll" _
    Alias "WNetAddConnectionA" ( _
        ByVal lpszNetPath As String, _
        ByVal lpszPassword As String, _
        ByVal lpszLocalName As String) _
    As Long

Declare Function WNetAddConnection2 _
    Lib "mpr.dll" _
    Alias "WNetAddConnection2A" ( _
        lpNetResource As NETRESOURCE, ByVal lpPassword As String, _
        ByVal lpUserName As String, ByVal dwFlags As Long) _
    As Long

Declare Function WNetCancelConnection _
    Lib "mpr.dll" _
    Alias "WNetCancelConnectionA" ( _
        ByVal lpszName As String, _
        ByVal bForce As Long) _
    As Long

Declare Function WNetCancelConnection2 _
    Lib "mpr.dll" _
    Alias "WNetCancelConnection2A" ( _
        ByVal lpName As String, _
        ByVal dwFlags As Long, _
        ByVal fForce As Long) _
    As Long

Declare Function WNetCloseEnum _
    Lib "mpr.dll" ( _
        ByVal hEnum As Long) _
    As Long

Declare Function WNetConnectionDialog _
    Lib "mpr.dll" ( _
        ByVal hWnd As Long, _
        ByVal dwType As Long) _
    As Long

Declare Function WNetDisconnectDialog _
    Lib "mpr.dll" ( _
        ByVal hWnd As Long, _
        ByVal dwType As Long) _
    As Long

Declare Function WNetEnumResource _
    Lib "mpr.dll" _
    Alias "WNetEnumResourceA" ( _
        ByVal hEnum As Long, lpcCount As Long, _
        lpBuffer As Any, lpBufferSize As Long) _
    As Long

Declare Function WNetGetConnection _
    Lib "mpr.dll" _
    Alias "WNetGetConnectionA" ( _
        ByVal lpszLocalName As String, _
        ByVal lpszRemoteName As String, _
        cbRemoteName As Long) _
    As Long

Declare Function WNetGetLastError _
    Lib "mpr.dll" _
    Alias "WNetGetLastErrorA" ( _
        lpError As Long, _
        ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, _
        ByVal lpNameBuf As String, ByVal nNameBufSize As Long) _
    As Long

Declare Function WNetGetUniversalName _
    Lib "mpr" _
    Alias "WNetGetUniversalNameA" ( _
        ByVal lpLocalPath As String, ByVal dwInfoLevel As Long, _
        lpBuffer As Any, lpBufferSize As Long) _
    As Long

Declare Function NetGetUser _
    Lib "mpr.dll" _
    Alias "WNetGetUserA" ( _
        ByVal lpName As String, _
        ByVal lpUserName As String, _
        lpnLength As Long) _
    As Long

Declare Function WNetGetUser _
    Lib "mpr" _
    Alias "WNetGetUserA" ( _
        ByVal lpName As String, _
        ByVal lpUserName As String, _
        lpnLength As Long) _
    As Long

Declare Function WNetOpenEnum _
    Lib "mpr.dll" _
    Alias "WNetOpenEnumA" ( _
        ByVal dwScope As Long, ByVal dwType As Long, _
        ByVal dwUsage As Long, lpNetResource As NETRESOURCE, _
        lphEnum As Long) _
    As Long


