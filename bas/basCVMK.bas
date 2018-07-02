Attribute VB_Name = "basCVMK"
Option Explicit

' Substitutes for the old CV* and MK* functions
' from qbasic.

Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
            hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
            
Public Function CVI(s As String) As Integer
   Dim i As Integer
   
   If Len(s) <> 2 Then
      Err.Raise 1000, "CVI", "Invalid string argument"
   Else
      CopyMemory i, ByVal s, 2
   End If
   
   CVI = i
End Function
Public Function CVL(s As String) As Long
   Dim i As Long
   
   If Len(s) <> 4 Then
      Err.Raise 1000, "CVL", "Invalid string argument"
   Else
      CopyMemory i, ByVal s, 4
   End If
   
   CVL = i
End Function
Public Function CVD(s As String) As Double
   Dim i As Double
   
   If Len(s) <> 8 Then
      Err.Raise 1000, "CVD", "Invalid string argument"
   Else
      CopyMemory i, ByVal s, 8
   End If
   
   CVD = i
End Function
Public Function CVS(s As String) As Single
   Dim i As Single
   
   If Len(s) <> 4 Then
      Err.Raise 1000, "CVS", "Invalid string argument"
   Else
      CopyMemory i, ByVal s, 4
   End If
   
   CVS = i
End Function
Public Function MKI(ByVal i As Integer) As String
    Dim s As String
    
    s = String(2, 0)
    CopyMemory ByVal s, i, 2
    
    MKI = s
End Function

Public Function MKL(ByVal i As Long) As String
    Dim s As String
    
    s = String(4, 0)
    CopyMemory ByVal s, i, 4
    
    MKL = s
End Function
Public Function MKS(ByVal i As Double) As String
    Dim s As String
    
    s = String(4, 0)
    CopyMemory ByVal s, i, 4
    
    MKS = s
End Function

Public Function MKD(ByVal i As Double) As String
    Dim s As String
    
    s = String(8, 0)
    CopyMemory ByVal s, i, 8
    
    MKD = s
End Function
