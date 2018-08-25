VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form pop3session 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pop3 Session"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   45
      Pattern         =   "*.txt"
      TabIndex        =   0
      Top             =   0
      Width           =   1170
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   2790
      Top             =   915
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "pop3session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public inbuf As String
Public outbuf As String

Public timestamp As String
Public username As String
Public password As String
Public transactionstate As Boolean
Public deletedmsgs As New Dictionary

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    Dim q As String, ts As TextStream
    ws.getdata q
    inbuf = inbuf & q
    If InStr(1, inbuf, vbCrLf) = 0 Then Exit Sub
    On Error Resume Next
    K = Mid(inbuf, 1, InStr(1, inbuf, " ") - 1)
    v = Mid(inbuf, InStr(1, inbuf, " ") + 1)
    inbuf = Mid(inbuf, InStr(1, inbuf, vbCrLf) + 2)
    If Right(v, 2) = vbCrLf Then v = Mid(v, 1, Len(v) - 2)
    If K = "" And v <> "" Then K = v: v = ""
    On Error GoTo 0
    Select Case transactionstate
    Case False
        Select Case UCase(K)
        Case "QUIT"
            outbuf = outbuf & "+OK cya" & vbCrLf
            ws.Close
            Me.Hide
            Unload Me
            Exit Sub
        Case "USER"
            username = v
            If username <> "" And fso.FolderExists(fso.BuildPath(subfolder(""), username)) Then
                outbuf = outbuf & "+OK good username" & vbCrLf
            Else
                outbuf = outbuf & "-ERR bad username" & vbCrLf
            End If
        Case "PASS"
            If LCase(getaccountinfo(username, "pw")) = LCase(v) Then
                password = v
                transactionstate = True
                outbuf = outbuf & "+OK good login" & vbCrLf
                File1.Path = subfolder(username)
            Else
                outbuf = outbuf & "-ERR negative login, please try again" & vbCrLf
                username = ""
                password = ""
            End If
        Case "APOP"
            digest = Mid(v, InStr(1, v, " ") + 1)
            v = Mid(v, 1, InStr(1, v, " ") - 1)
            username = v
            If LCase(digest) = LCase(checksum(timestamp & getaccountinfo(username, "pw"))) Then
                transactionstate = True
                outbuf = outbuf & "+OK secure login fine" & vbCrLf
            Else
                outbuf = outbud & "-ERR Error in secure login" & vbCrLf
            End If
        Case Else
            outbuf = outbuf & "-ERR Bad command at authenticate stage" & vbCrLf
        End Select
    Case True 'We know who we are now, implement the real jucy stuff
        Select Case UCase(K)
        Case "NOOP"
            outbuf = outbuf & "+OK I pong your ping" & vbCrLf
        Case "DELE"
            deletedmsgs(v) = True
            outbuf = outbuf & "+OK msg will deleted on session end" & vbCrLf
        Case "RSET"
            deletedmsgs.RemoveAll
            outbuf = outbuf & "+OK all msgs are no longer deleted" & vbCrLf
        Case "UIDL"
            If v = "" Then
                outbuf = outbuf & "+OK" & vbCrLf
                For a = 1 To File1.ListCount - 1
                    outbuf = outbuf & a & " " & Mid(File1.List(a), 1, Len(File1.List(a)) - 4) & vbCrLf
                Next a
                outbuf = outbuf & "." & vbCrLf
            Else
                outbuf = outbuf & "+OK " & v & " " & Mid(File1.List(v), 1, Len(File1.List(v)) - 4) & vbCrLf
            End If
        Case "LIST"
            If v = "" Then
                outbuf = outbuf & "+OK" & vbCrLf
                For a = 1 To File1.ListCount - 1
                    outbuf = outbuf & a & " " & getmailsize(fso.BuildPath(File1.Path, File1.List(a))) & vbCrLf
                Next a
                outbuf = outbuf & "." & vbCrLf
            Else
                outbuf = outbuf & "+OK " & v & " " & getmailsize(fso.BuildPath(File1.Path, File1.List(v))) & vbCrLf
            End If
        Case "STAT"
            Sum = 0
            For a = 1 To File1.ListCount - 1
                Sum = Sum + getmailsize(fso.BuildPath(File1.Path, File1.List(a)))
            Next a
            outbuf = outbuf & "+OK " & File1.ListCount - 1 & " " & Sum & vbCrLf
        Case "RETR"
            Set ts = fso.OpenTextFile(fso.BuildPath(File1.Path, File1.List(v)))
            ts.SkipLine
            ts.SkipLine
            ts.SkipLine
            ts.SkipLine
            outbuf = outbuf & "+OK" & vbCrLf & ts.ReadAll & vbCrLf & "." & vbCrLf
            ts.Close
        Case "TOP"
            numlines = Mid(v, InStr(1, v, " ") + 1)
            v = Mid(v, 1, InStr(1, v, " ") - 1)
            Set ts = fso.OpenTextFile(fso.BuildPath(File1.Path, File1.List(v)))
            ts.SkipLine
            ts.SkipLine
            ts.SkipLine
            ts.SkipLine
            outbuf = outbuf & "+OK" & vbCrLf
keepgoing:
            z = ts.ReadLine
            If z <> "" Then outbuf = outbuf & z & vbCrLf: GoTo keepgoing
            outbuf = outbuf & vbCrLf
            On Error Resume Next
            For a = 1 To numlines
                outbuf = outbuf & ts.ReadLine & vbCrLf
            Next a
            On Error GoTo 0
            outbuf = outbuf & "." & vbCrLf
            ts.Close
        Case "QUIT"
            outbuf = outbuf & "+OK cya" & vbCrLf
            For Each a In deletedmsgs.Keys
                fso.DeleteFile fso.BuildPath(File1.Path, File1.List(a)), True
            Next
        Case Else
            outbuf = outbuf & "-ERR bad command" & vbCrLf
        End Select
    End Select
    
    ws.SendData outbuf
    outbuf = ""
End Sub

Private Function getmailsize(FileName As String) As Long
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(FileName)
    ts.SkipLine
    ts.SkipLine
    ts.SkipLine
    ts.SkipLine
    content = ts.ReadAll
    getmailsize = Len(content)
    ts.Close
End Function

Private Sub ws_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ws.Close
    Me.Hide
    Unload Me
End Sub

Private Function previewtopline() As String
    a = InStr(1, inbuf, vbCrLf)
    If a > 0 Then
        previewtopline = Mid(inbuf, 1, a - 1)
    End If
End Function

Private Function pulltopline() As String
    a = InStr(1, inbuf, vbCrLf)
    If a > 0 Then
        pulltop = Mid(inbuf, 1, a - 1)
        inbuf = Mid(inbuf, a + 2)
    End If
End Function
