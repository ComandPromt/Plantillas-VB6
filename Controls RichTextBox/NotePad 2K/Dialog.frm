VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encryption & Decryption"
   ClientHeight    =   3885
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton DontUse 
      Caption         =   "Dont Use 3rd Encryption"
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   3240
      Width           =   2055
   End
   Begin VB.OptionButton Use 
      Caption         =   "Use 3rd Encryption too"
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2520
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.OptionButton UseOnly 
      Caption         =   "Use 3rd Encryption Only"
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Encrypt in Format"
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      ToolTipText     =   "Encrypt colors, fonts, etc.  Recommended if Copyed then to paste as text"
      Top             =   3600
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "DONE"
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Height          =   495
      Left            =   4080
      TabIndex        =   12
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Encrypt2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Encrypting Code #2"
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox Encrypt1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Encrypting Code #1"
      Top             =   720
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "3rd Encryption"
      Height          =   2175
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   3735
      Begin VB.ComboBox Super 
         Height          =   315
         ItemData        =   "Dialog.frx":0442
         Left            =   2040
         List            =   "Dialog.frx":044C
         TabIndex        =   8
         Text            =   "True"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Encrypt7 
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         ToolTipText     =   "Encrypt Code 7"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Encrypt4 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Encrypt Code 4"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Encrypt6 
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         ToolTipText     =   "Encrypt Code 6"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Encrypt5 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Encrypt Code 5"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Encrypt3 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Encrypt Code 3"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Super Encrypt"
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Encrypt Code 7"
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Encrypt Code 6"
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Encrypt Code 4"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Encrypt Code 5"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Encrypt Code 3"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No bigger than 130"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Status of Encryption"
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Number can be up to 10000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Encrypted, Decrypted As String
Private Sub Command1_Click()
    On Error GoTo Help
    Dim EncryptThis As String, SuperEncrypt As Boolean
    If Super.Text = "True" Then
        SuperEncrypt = True
    Else
        SuperEncrypt = False
    End If
    If UseOnly.Value = True Then
        Screen.MousePointer = vbHourglass
        If Check2.Value = 1 Then
            EncryptThis = fMainForm.Text1.TextRTF
        Else
            EncryptThis = fMainForm.Text1.Text
        End If
        Call Encryption1(Val(Encrypt3.Text), Val(Encrypt4.Text), Val(Encrypt5.Text), Val(Encrypt6.Text), Val(Encrypt7.Text), SuperEncrypt, EncryptThis)
        fMainForm.Text1.Text = Encrypted
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Val(Encrypt1.Text) > 130 Then
        Label1.Caption = "This number is to big"
        Exit Sub
    End If
    If Val(Encrypt1.Text) > 10000 Then
        Label2.Caption = "This number is to big"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Label1.Caption = "No Bigger than 130"
    Label2.Caption = "Number can be up to 10000"
    Label3.Caption = "Status"
    Dim letter, Encrypt, frmText, Encryption As String
    Encrypt = ""
    If Check2.Value = 1 Then
        frmText = fMainForm.Text1.TextRTF
    Else
        frmText = fMainForm.Text1.Text
    End If
    Dim i As Long
    For i = 1 To Len(frmText)
        letter = Mid(frmText, i, 1)
        Encrypt = Encrypt + Chr(Asc(letter) + Val(Encrypt1))
    Next i
    Encryption = ""
    For i = 1 To Len(Encrypt)
        letter = Mid(Encrypt, i, 1)
        Encryption = Encryption & (Asc(letter) Xor Val(Encrypt2))
        Encryption = Encryption + " "
        DoEvents
    Next i
    If Use.Value = True Then
        Call Encryption1(Val(Encrypt3.Text), Val(Encrypt4.Text), Val(Encrypt5.Text), Val(Encrypt6.Text), Val(Encrypt7.Text), SuperEncrypt, Encryption)
        Encrypt = Encrypted
    Else
        Encrypt = Encryption
    End If
    fMainForm.Text1.Text = Encrypt
    Screen.MousePointer = vbDefault
    Exit Sub
Help:
    Beep
    Screen.MousePointer = vbDefault
    Label3.Caption = "Error: " + Err.Description
End Sub

Private Sub Command2_Click()
    On Error GoTo Help
    Dim EncryptThis As String, SuperEncrypt As Boolean
    If Super.Text = "True" Then
        SuperEncrypt = True
    Else
        SuperEncrypt = False
    End If
    If UseOnly.Value = True Then
        Screen.MousePointer = vbHourglass

        Call Decryption1(Val(Encrypt3.Text), Val(Encrypt4.Text), Val(Encrypt5.Text), Val(Encrypt6.Text), Val(Encrypt7.Text), SuperEncrypt, fMainForm.Text1.Text)
        If Check2.Value = 1 Then
            fMainForm.Text1.TextRTF = Decrypted
        Else
            fMainForm.Text1.Text = Decrypted
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Val(Encrypt1.Text) > 130 Then
        Label1.Caption = "This number is to big"
        Exit Sub
    End If
    If Val(Encrypt1.Text) > 10000 Then
        Label2.Caption = "This number is to big"
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    Label1.Caption = "No Bigger than 130"
    Label2.Caption = "Number can be up to 10000"
    Label3.Caption = "Status"
    Dim Decryption, Num, letter2, Decrypt, letter As String
    Dim Number As Long
    Decrypt = ""
    Decryption = ""
    Dim i, Numb As Integer
    Numb = 1
    Decryption = ""
    If Use.Value = True Then
        Call Decryption1(Val(Encrypt3.Text), Val(Encrypt4.Text), Val(Encrypt5.Text), Val(Encrypt6.Text), Val(Encrypt7.Text), SuperEncrypt, fMainForm.Text1.Text)
        Decryption = Decrypted
    Else
        Decryption = fMainForm.Text1.Text
    End If
    For i = 1 To Len(Decryption)
        DoEvents
        letter = ""
        Num = ""
        Do Until letter = " "
            letter = Mid(Decryption, Numb, 1)
            Num = Num & letter
            Numb = Numb + 1
        Loop
        letter2 = Chr(Val(Trim(Num)) Xor Val(Encrypt2))
        Decrypt = Decrypt + letter2
        If Numb >= Len(Decryption) Then Exit For
    Next i
    
    Dim Finish As String
    For i = 1 To Len(Decrypt)
        letter = Mid(Decrypt, i, 1)
        Finish = Finish + Chr(Asc(letter) - Val(Encrypt1))
    Next i
    If Check2.Value = 1 Then
        fMainForm.Text1.TextRTF = Finish
    Else
        fMainForm.Text1.Text = Finish
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Help:
    Beep
    Screen.MousePointer = vbDefault
    Label3.Caption = "Error: " + Err.Description
End Sub

Private Sub Form_Initialize()
    Label3.Caption = "Status"
    Encrypt1.Text = ""
    Encrypt2.Text = ""
End Sub

Private Sub Form_Load()
    Label3.Caption = "Status"
    Encrypt1.Text = ""
    Encrypt2.Text = ""
    Dialog.Check2.Value = GetSetting(App.Title, "Check", "Encryption Format", 1)
    Dialog.Use.Value = GetSetting(App.Title, "Option", "Use", False)
    Dialog.UseOnly.Value = GetSetting(App.Title, "Option", "Useonly", False)
    Dialog.DontUse.Value = GetSetting(App.Title, "Option", "DontUse", True)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Check", "Encryption Format", Dialog.Check2.Value
    SaveSetting App.Title, "Option", "Use", Dialog.Use.Value
    SaveSetting App.Title, "Option", "Useonly", Dialog.UseOnly.Value
    SaveSetting App.Title, "Option", "DontUse", Dialog.DontUse.Value
End Sub

Private Sub OKButton_Click()
    Dialog.Hide
End Sub





'3rd Encryption Takin' Planet Source Code
Function Encryption1(start As Long, diff As Long, beta As Integer, alpha As Integer, times As Long, SuperEncrypt As Boolean, Text As String)
    'Encrypt characters
    On Error GoTo error
    Dim i As Long
    Dim curkey As Long
    Dim m As Long
    Dim endstr As String
    Dim Text2 As String
    Dim lesser As Double
    Dim larger As Double
    Dim SuperE As Boolean
    Dim a As Long
    SuperE = SuperEncrypt


    If diff > 500 Then
        diff = 500
    ElseIf diff < 1 Then
        diff = 1
    End If


    If times > 100 Then
        times = 100
    ElseIf times < 1 Then
        times = 1
    End If


    If start > 255 Then
        start = 255
    ElseIf start < 1 Then
        start = 1
    End If


    If beta > 5 Then
        beta = 5
    ElseIf beta < 1 Then
        beta = 1
    End If


    If alpha > 5 Then
        alpha = 5
    ElseIf alpha < 1 Then
        alpha = 1
    End If
    curkey = start
    curkey = (curkey * alpha) / beta


    If SuperE = True Then


        If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
            curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
        Else
            curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
        End If
        curkey = SuperEE(curkey, beta, alpha, beta)
    End If


    If curkey > 255 Then
        curkey = 255 - (curkey / 255)
    ElseIf curkey < 0 Then
        curkey = 0 - (curkey / 255)
    End If


    For a = 1 To times
        DoEvents

        For i = 1 To Len(Text)


            If 255 - curkey > curkey Then
                larger = 255 - curkey
                lesser = curkey
            Else
                larger = curkey
                lesser = 255 - curkey
            End If


            If Asc(Mid$(Text, i, 1)) <= lesser Then
                m = Asc(Mid$(Text, i, 1)) + (larger - 1)
                endstr = endstr + Chr$(m)
            Else
                m = Asc(Mid$(Text, i, 1)) - lesser
                endstr = endstr + Chr$(m)
            End If
            curkey = curkey + diff

            DoEvents
            If curkey > 255 Then
                curkey = curkey - 255
            End If
            curkey = (curkey * alpha) / beta


            If SuperE = True Then


                If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
                    curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
                Else
                    curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
                End If
                curkey = SuperEE(curkey, beta, alpha, beta)
            End If
            beta = beta + (2 * diff)
            alpha = alpha + diff


            If beta > 5 Then
                beta = 1
            End If


            If alpha > 5 Then
                alpha = 1
            End If


            If curkey > 255 Then
                curkey = 255 - (curkey / 255)
            ElseIf curkey < 0 Then
                curkey = 0 - (curkey / 255)
            End If


            If diff > 500 Then
                diff = 1
            Else
                diff = diff + diff
            End If
        Next i
        Text2 = ""
        Text2 = endstr
        endstr = ""
    Next a
    Encrypted = Text2
    Exit Function
error:
    Beep
    Label3.Caption = "Error: " + Err.Description
End Function


Function Decryption1(start As Long, diff As Long, beta As Integer, alpha As Integer, times As Long, SuperEncrypt As Boolean, Text As String)
    'Decrypt characters
    On Error GoTo error
    Dim i As Long
    Dim curkey As Long
    Dim m As Long
    Dim endstr As String
    Dim Text2 As String
    Dim lesser As Double
    Dim larger As Double
    Dim SuperE As Boolean
    Dim a As Long
    SuperE = SuperEncrypt


    If diff > 500 Then
        diff = 500
    ElseIf diff < 1 Then
        diff = 1
    End If


    If times > 100 Then
        times = 100
    ElseIf times < 1 Then
        times = 1
    End If


    If start > 255 Then
        start = 255
    ElseIf start < 1 Then
        start = 1
    End If


    If beta > 5 Then
        beta = 5
    ElseIf beta < 1 Then
        beta = 1
    End If


    If alpha > 5 Then
        alpha = 5
    ElseIf alpha < 1 Then
        alpha = 1
    End If
    curkey = start
    curkey = (curkey * alpha) / beta


    If SuperE = True Then


        If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
            curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
        Else
            curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
        End If
        curkey = SuperEE(curkey, beta, alpha, beta)
    End If


    If curkey > 255 Then
        curkey = 255 - (curkey / 255)
    ElseIf curkey < 0 Then
        curkey = 0 - (curkey / 255)
    End If


    For a = 1 To times

        DoEvents
        For i = 1 To Len(Text)


            If 255 - curkey > curkey Then
                larger = 255 - curkey
                lesser = curkey
            Else
                larger = curkey
                lesser = 255 - curkey
            End If


            If Asc(Mid$(Text, i, 1)) >= larger Then
                m = Asc(Mid$(Text, i, 1)) - (larger - 1)
                endstr = endstr + Chr$(m)
            Else
                m = Asc(Mid$(Text, i, 1)) + lesser
                endstr = endstr + Chr$(m)
            End If
            curkey = curkey + diff


            If curkey > 255 Then
                curkey = curkey - 255
            End If
            curkey = (curkey * alpha) / beta
            
            DoEvents

            If SuperE = True Then


                If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
                    curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
                Else
                    curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
                End If
                curkey = SuperEE(curkey, beta, alpha, beta)
            End If
            beta = beta + (2 * diff)
            alpha = alpha + diff


            If beta > 5 Then
                beta = 1
            End If


            If alpha > 5 Then
                alpha = 1
            End If


            If curkey > 255 Then
                curkey = 255 - (curkey / 255)
            ElseIf curkey < 0 Then
                curkey = 0 - (curkey / 255)
            End If


            If diff > 500 Then
                diff = 1
            Else
                diff = diff + diff
            End If
        Next i
        Text2 = ""
        Text2 = endstr
        endstr = ""
    Next a
    Decrypted = Text2
    Exit Function
error:
    Beep
    Label3.Caption = "Error: " + Err.Description
End Function

Private Function SuperEE(curkey As Long, beta As Integer, alpha As Integer, times As Integer)
    'For encryption: Change the current key
    '     around more
    On Error Resume Next
    curkey = (((curkey / times) - (beta + times)) * alpha) + ((beta / alpha) - times)


    If curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) < 1 Then
        curkey = (((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10)) * (0 - 1))
    Else
        curkey = ((curkey + beta) * alpha) - (((curkey - beta) + alpha) / ((beta - alpha) + 10))
    End If


    If beta - times = 0 Then
        curkey = ((curkey * alpha) + (beta * times))
    Else
        curkey = ((curkey * (beta - times)) + (beta - times))


        If curkey < 0 Then
            curkey = curkey + (alpha + beta)
        ElseIf curkey = 0 Then
            curkey = curkey + (alpha + times)
        Else
            curkey = curkey + (beta + times)
        End If
    End If
    SuperEE = curkey
End Function
