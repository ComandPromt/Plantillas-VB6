Attribute VB_Name = "MConvertSize"
Option Explicit

Private Declare Function StrFormatByteSize Lib _
    "shlwapi" Alias "StrFormatByteSizeA" (ByVal _
    dw As Long, ByVal pszBuf As String, ByRef _
    cchBuf As Long) As String

Public Function FormatKB(ByVal Amount As Long) _
    As String
    Dim Buffer As String
    Dim Result As String
    Buffer = Space$(255)
    Result = StrFormatByteSize(Amount, Buffer, _
        Len(Buffer))
    If InStr(Result, vbNullChar) > 1 Then
        FormatKB = Left$(Result, InStr(Result, _
            vbNullChar) - 1)
    End If
End Function


