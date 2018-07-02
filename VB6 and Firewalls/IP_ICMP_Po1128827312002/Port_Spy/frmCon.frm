VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCon 
   Caption         =   "Test Connection"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCon 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5040
      Top             =   1080
   End
   Begin MSComCtl2.Animation aniNET 
      Height          =   855
      Left            =   4800
      TabIndex        =   13
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
      _Version        =   393216
      FullWidth       =   49
      FullHeight      =   57
   End
   Begin VB.Frame fraCon2 
      Caption         =   ".:: Connected to: <Not Connected> ::."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   5775
      Begin VB.CheckBox chkEnable 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   4680
         TabIndex        =   24
         Top             =   960
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.Frame fraOS 
         Appearance      =   0  'Flat
         Caption         =   ".:: OS Emulation ::."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   4335
         Begin VB.OptionButton optEM 
            Caption         =   "Windows 95"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optEM 
            Caption         =   "Windows 98"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2280
            TabIndex        =   22
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optEM 
            Caption         =   "Windows 2000"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton optEM 
            Caption         =   "Windows XP"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   2280
            TabIndex        =   20
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.Frame fraClient 
         Appearance      =   0  'Flat
         Caption         =   ".:: Client Emulation ::."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   4335
         Begin VB.OptionButton optEM 
            Caption         =   "Mozilla 0.9.5 (web)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   2280
            TabIndex        =   18
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton optEM 
            Caption         =   "Opera 5.01 (web)"
            Enabled         =   0   'False
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton optEM 
            Caption         =   "Netscape 6.2 (web)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   2280
            TabIndex        =   16
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton optEM 
            Caption         =   "IE 6.0 (web)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   4680
         TabIndex        =   12
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdDisCon 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4680
         TabIndex        =   11
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4680
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtSend 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Send text:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraCon1 
      Caption         =   ".:: Connection ::. "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtPort 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtAddr 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "1 to 65536"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin MSWinsockLib.Winsock WS1 
      Left            =   4560
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' Copyright © 2001 Gregory Kirk. All rights reserved.
'
' You have a royalty-free right to use, modify, reproduce and distribute the
' Application Files (and/or any modified version) in any way you find useful,
' provided that you agree that Gregory Kirk has no warranty, obligations or
' liability for any Application Files.
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' Connection form:
'    The whole purpose of this is to see server headers if any or any data for
'    that matter sent by a server. When I include UDP, SMTP/POP3 this will be
'    made into a seperate program with trace route.
'-------------------------------------------------------------------------------

Option Explicit
Dim TimeOut As Integer 'Set to 30 seconds
Dim OS As Integer, CL As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : Form_Load
' Purpose    : Set default settings.
' Parameters : NA
' Return val : NA
' Algorithm  : aniNET.Open
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
TimeOut = 0
aniNET.Open App.Path & "\xpsearchinternet.avi" 'Not really neccessary but ey it looks good.
chkEnable_Click
With txtSend
    .Text = ""
    .Enabled = False
    .BackColor = &H8000000F
End With
cmdSend.Enabled = False
cmdDisCon.Enabled = False
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : Form_Unload
' Purpose    : Close open socket, stop and close aniNet, unload frmView.
' Parameters : Cancel As Integer(not used)
' Return val : NA
' Algorithm  : aniNET.AutoPlay; aniNET.Close
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
If WS1.State <> sckClosed Then WS1.Close
If aniNET.AutoPlay = True Then aniNET.AutoPlay = False
aniNET.Close
Unload frmView
Unload Me
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : cmdConnect_Click
' Purpose    : Try to connect to address, show data viewer.
' Parameters : NA
' Return val : NA
' Algorithm  : aniNET.AutoPlay; WS1.Connect
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdConnect_Click()
    aniNET.AutoPlay = True
    WS1.Connect txtAddr.Text, txtPort.Text
    fraCon2.Caption = ".:: Connected to: <Not Connected> ::."
    frmView.Show
    frmView.txtConView.Text = "<<< Trying to connect to: " & txtAddr.Text & ":" & txtPort.Text & " >>>"
    cmdDisCon.Enabled = True
    cmdDisCon.Caption = "&Abort"
    tmrCon.Enabled = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : tmrCon_Timer
' Purpose    : Timer for connection, times out after 30 seconds.
' Parameters : NA
' Return val : NA
' Algorithm  : Calls ConSuccess
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrCon_Timer()
TimeOut = TimeOut + 1
    If WS1.State = sckConnected Then
        TimeOut = 0
        ConSuccess
        Exit Sub
    ElseIf TimeOut >= 15 Then
        MsgBox "Connection timed out after 30 seconds...", vbExclamation, ".:: Connection Error ::."
        Unload frmView
        If tmrCon.Enabled = True Then tmrCon.Enabled = False 'avoid getting stuck in a nasty msgbox popup loop!
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : ConSuccess
' Purpose    : Connection established, send data to be sent.
' Parameters : NA
' Return val : NA
' Algorithm  : aniNET.AutoPlay; WS1.SendData
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ConSuccess()
    tmrCon.Enabled = False
    fraCon2.Caption = "Connected to: " & txtAddr.Text & ":" & txtPort.Text
    aniNET.AutoPlay = False
    
frmView.txtConView.Text = frmView.txtConView.Text & vbCrLf & "<<< ready to send/recv >>>"
If chkEnable.Value = 1 Then
    Dim sndData As String
    sndData = head(OS, CL)
    frmView.txtConView.Text = frmView.txtConView.Text & vbCrLf & "<<< " & sndData & " >>>"
    WS1.SendData sndData
    DoEvents
Else
    With txtSend
        .Text = ""
        .Enabled = True
        .BackColor = vbWhite
    End With
    cmdSend.Enabled = True
End If
cmdDisCon.Caption = "Disconnect"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : chkEnable_Click
' Purpose    : Set emulation parameters.
' Parameters : NA
' Return val : NA
' Algorithm  : NA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkEnable_Click()
Dim i As Integer, tmp As String
If chkEnable.Value = 1 Then
    fraOS.Enabled = True
    fraClient.Enabled = True
    For i = 0 To 7
        optEM(i).Enabled = True
    Next
    With txtSend
        .Text = "<Internal data to be sent>"
        .Enabled = False
        .BackColor = &H8000000F
    End With
    cmdSend.Enabled = False
    OS = 0
    CL = 4
Else
    fraOS.Enabled = False
    fraClient.Enabled = False
    For i = 0 To 7
        optEM(i).Enabled = False
    Next
    If cmdDisCon.Enabled = True Then
        With txtSend
            .Text = ""
            .Enabled = True
            .BackColor = vbWhite
        End With
        cmdSend.Enabled = True
    End If
End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name             : optEM_Click
' Purpose          : Set emulation parameters.
' Parameters       : Index As Integer
' Return val       : NA
' Algorithm        : NA
' Optimization note: Given an array of items and only one can be true at a time, then
' it is faster to test for the value itself than the value of every item in the array.
' If an item prior to the last is true then there is no need to check the rest of them,
' therefore exit loop. If you have an array of several hundred values this can seriously
' save time.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub optEM_Click(Index As Integer)
Dim i As Integer
For i = 0 To 3
    Select Case True
        Case optEM(i).Value: OS = i: Exit For
    End Select
Next
For i = 4 To 7
    Select Case True
        Case optEM(i).Value: CL = i: Exit For
    End Select
Next
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : cmdSend_Click
' Purpose    : Send data to be sent.
' Parameters : NA
' Return val : NA
' Algorithm  : WS1.SendData
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSend_Click()
If WS1.State <> sckConnected Or WS1.State <> sckOpen Then
    cmdDisCon_Click
    MsgBox "Conection refused...", vbExclamation, ".:: Error Sending Data ::."
    Exit Sub
End If
frmView.txtConView.Text = frmView.txtConView.Text & vbCrLf & "// " & txtSend.Text & " \\"
WS1.SendData txtSend.Text & Chr$(10)
DoEvents
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : cmdDisCon_Click
' Purpose    : Close socket and set everything to default state.
' Parameters : NA
' Return val : NA
' Algorithm  : aniNET.AutoPlay
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDisCon_Click()
If WS1.State <> sckClosed Then WS1.Close
aniNET.AutoPlay = False
With txtSend
    .Enabled = False
    .BackColor = &H8000000F
End With
cmdSend.Enabled = False
cmdDisCon.Enabled = False
fraCon2.Caption = ".:: Connected to: <Not Connected> ::."
Unload frmView
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : cmdExit_Click
' Purpose    : Exit frmCon.
' Parameters : NA
' Return val : NA
' Algorithm  : NA
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdExit_Click()
Unload Me
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : WS1_DataArrival
' Purpose    : Get data sent by the server.
' Parameters : ByVal bytesTotal As Long
' Return val : strData
' Algorithm  : WS1.GetData
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub WS1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
    WS1.GetData strData
    frmView.txtConView.Text = frmView.txtConView.Text & vbCrLf & strData
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name       : WS1_Error
' Purpose    : This is raised when the server rejects a connection attempt or connection cannot be established.
' Parameters : ByVal Number As Integer, Description As String, ByVal Scode As Long
'              ByVal Source As String, ByVal HelpFile As String
'              ByVal HelpContext As Long, CancelDisplay As Boolean
' Return val : NA
' Algorithm  : WS1.Close
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub WS1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If aniNET.AutoPlay = True Then aniNET.AutoPlay = False
WS1.Close
With txtSend
    .Text = ""
    .Enabled = False
    .BackColor = &H8000000F
End With
If tmrCon.Enabled = True Then tmrCon.Enabled = False 'Avoid getting stuck in a msgbox popup loop.
cmdSend.Enabled = False
cmdDisCon.Enabled = False
MsgBox "Conection refused...", vbExclamation, ".:: Error Sending Data ::."
End Sub
