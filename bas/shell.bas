Attribute VB_Name = "Module2"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long





Public Function ShellFile(Program As String, Optional RunStyle As Long = vbNormalFocus, Optional ByVal WorkDir As Variant) As Long

    Dim FirstSpace As Integer, Slash As Integer
On Error GoTo 98

    If Left(Program, 1) = """" Then
        FirstSpace = InStr(2, Program, """")


        If FirstSpace <> 0 Then
            Program = Mid(Program, 2, FirstSpace - 2) & Mid(Program, FirstSpace + 1)
            FirstSpace = FirstSpace - 1
        End If

    Else
        FirstSpace = InStr(Program, " ")
    End If

    If FirstSpace = 0 Then FirstSpace = Len(Program) + 1


    If IsMissing(WorkDir) Then


        For Slash = FirstSpace - 1 To 1 Step -1
            If Mid(Program, Slash, 1) = "\" Then Exit For
        Next



        If Slash = 0 Then
            WorkDir = CurDir
        ElseIf Slash = 1 Or Mid(Program, Slash - 1, 1) = ":" Then
            WorkDir = Left(Program, Slash)
        Else
            WorkDir = Left(Program, Slash - 1)
        End If

    End If

    ShellFile = ShellExecute(0, vbNullString, _
    Left(Program, FirstSpace - 1), LTrim(Mid(Program, FirstSpace)), _
    WorkDir, RunStyle)

    If ShellFile < 32 Then VBA.Shell Program, RunStyle 'To raise Error
Form1.Send "Shelled File (" & Program & ")"

Exit Function
98
Form1.Send "Shell Error " & "(" & Err.Description & ")"
End Function




