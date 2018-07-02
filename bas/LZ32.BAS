Attribute VB_Name = "basLZ32"
' Module Name               basLZ32
' Module Description        Module which contains ALL Declarations
'                           for API calls in the LZ32 Library
'
Option Explicit

Declare Function CopyLZFile _
    Lib "lz32" ( _
        ByVal n1 As Long, _
        ByVal n2 As Long) _
    As Long
    
Declare Function GetExpandedName _
    Lib "lz32.dll" _
    Alias "GetExpandedNameA" ( _
        ByVal lpszSource As String, _
        ByVal lpszBuffer As String) _
    As Long
    
Declare Sub LZClose _
    Lib "lz32.dll" ( _
        ByVal hfFile As Long)


