Attribute VB_Name = "Module1"
Option Explicit

Declare Function BYTE2BIN Lib "BinWorks.dll" (num As Byte) As String
Declare Function INT2BIN Lib "BinWorks.dll" (num As Integer) As String
Declare Function LONG2BIN Lib "BinWorks.dll" (num As Long) As String

Declare Function BIN2BYTE Lib "BinWorks.dll" (ByVal BinStr As String) As Byte
Declare Function BIN2INT Lib "BinWorks.dll" (ByVal BinStr As String) As Integer
Declare Function BIN2LONG Lib "BinWorks.dll" (ByVal BinStr As String) As Long

Declare Function B_LSHIFT Lib "BinWorks.dll" (num As Byte, amt As Byte) As Byte
Declare Function B_RSHIFT Lib "BinWorks.dll" (num As Byte, amt As Byte) As Byte
Declare Function I_LSHIFT Lib "BinWorks.dll" (num As Integer, amt As Byte) As Integer
Declare Function I_RSHIFT Lib "BinWorks.dll" (num As Integer, amt As Byte) As Integer
Declare Function L_LSHIFT Lib "BinWorks.dll" (num As Long, amt As Byte) As Long
Declare Function L_RSHIFT Lib "BinWorks.dll" (num As Long, amt As Byte) As Long

Declare Function B_LROTATE Lib "BinWorks.dll" (num As Byte, amt As Byte) As Byte
Declare Function B_RROTATE Lib "BinWorks.dll" (num As Byte, amt As Byte) As Byte
Declare Function I_LROTATE Lib "BinWorks.dll" (num As Integer, amt As Byte) As Integer
Declare Function I_RROTATE Lib "BinWorks.dll" (num As Integer, amt As Byte) As Integer
Declare Function L_LROTATE Lib "BinWorks.dll" (num As Long, amt As Byte) As Long
Declare Function L_RROTATE Lib "BinWorks.dll" (num As Long, amt As Byte) As Long



