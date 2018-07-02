Attribute VB_Name = "modSelfExtract"
'If you are going to use this in a app, you must
'first contact me at aandrei@hades.ro, and you
'have to credit me on the application's box, and/or
'about box

Public TheFile As String 'You can get the file
                         'inside the exe from
                         'anywhere by accesing
                         'this variable, after
                         'you called the SelfExtract
                         'sub.

Sub SelfExtract()

On Error GoTo ErrHandler:

Dim Size As String
Dim iFreeFile As Integer

iFreeFile = FreeFile

'If Dir(App.Path & "\" & App.EXEName & ".exe") = "" Then
'    MsgBox "You must compile the project to a EXE before running it!", vbCritical, "Not available from IDE"
'    End
'End If

Open "c:\windows\desktop\selfextract.exe" For Binary As iFreeFile
    'get the size of the file
    Seek #iFreeFile, LOF(iFreeFile) - 11
    Size = String(10, Chr(0))
    Get iFreeFile, , Size
    Size = CCur(Size)           'convert to currency, to avoid
                                'overflow if the file is bigger
                                'than 2,147,483,648 bytes (just
                                'in case :))
    'ok... now get the file
    Seek #iFreeFile, LOF(iFreeFile) - 11 - CCur(Size)
    TheFile = String(Size, Chr(0))
    Get iFreeFile, , TheFile            'TheFile now contains the file inside your exe
Close iFreeFile

Exit Sub

ErrHandler:

Result = MsgBox("Error #" & Err.Number & _
    " while trying to extract the file." _
    & vbCrLf & "Description: " & Err.Description _
    & vbCrLf & "How do you want to continue?", _
    vbAbortRetryIgnore + vbExclamation, "Error")

If Result = vbRetry Then
    Resume
ElseIf Result = vbIgnore Then
    Resume Next
ElseIf Result = vbAbort Then
    End
End If

End Sub


