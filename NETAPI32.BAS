Attribute VB_Name = "basNetAPI32"
' Module Name               basNetAPI32
' Module Description        Module which contains ALL Declarations
'                           for API calls in the NetAPI32 Library
'
Option Explicit

Declare Function Netbios _
    Lib "netapi32.dll" ( _
        pncb As NCB) _
    As Byte

