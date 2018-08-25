VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmBS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Billing Summary"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Print 3"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print 2"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print 1"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   8760
      Width           =   975
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   8775
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   15478
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmBS.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   255
      Left            =   10800
      TabIndex        =   0
      Top             =   8760
      Width           =   975
   End
End
Attribute VB_Name = "frmBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Text1.SelPrint Printer.hDC
End Sub

Private Sub Command3_Click()
    Text1.SelPrint Printer.hDC
    Text1.SelPrint Printer.hDC
End Sub

Private Sub Command4_Click()
    Text1.SelPrint Printer.hDC
    Text1.SelPrint Printer.hDC
    Text1.SelPrint Printer.hDC
End Sub

Private Sub Form_Load()
    Dim Data As String
    Dim OB As Double
    Dim Total As Double
    Dim CC As Double
    
    OB = 0
    Total = 0
    CC = 0

    Text1.LoadFile App.Path & "\Billing Summary Default.rtf"
    
    With frmMain
        .Accounts.MoveFirst
        Text1.Text = Text1.Text & vbCrLf
        Text1.Text = Text1.Text & "Airline" & vbTab & "Invoice #" & vbTab & "Current" & vbTab & "Outstanding" & vbTab & "Total"
        Text1.Text = Text1.Text & vbCrLf
        Do Until .Accounts.EOF = True
            If .Accounts("Edited") = "NO" Then GoTo Skipy
            If Len(.Accounts!Airline) < 23 Then
                Data = (.Accounts!Airline & Space(23 - Len(.Accounts!Airline))) & vbTab & .Accounts!InvoiceNum & vbTab & Format(.Accounts!CurrentCharges, "$0.00") & vbTab & Format(.Accounts!OutstandingBalance, "$0.00") & vbTab & Format(.Accounts!Total, "$0.00")
            Else
                Data = Left(.Accounts!Airline, 23) & vbTab & .Accounts!InvoiceNum & vbTab & Format(.Accounts!CurrentCharges, "$0.00") & vbTab & Format(.Accounts!OutstandingBalance, "$0.00") & vbTab & Format(.Accounts!Total, "$0.00")
            End If
            OB = OB + .Accounts!OutstandingBalance
            Total = Total + .Accounts!Total
            CC = CC + .Accounts!CurrentCharges
            Text1.Text = Text1.Text & Data & vbCrLf
Skipy:
            .Accounts.MoveNext
        Loop
        Text1.Text = Text1.Text & vbCrLf & "Total:" & vbTab & vbTab & Format(CC, "$#0.00") & vbTab & Format(OB, "$#0.00") & vbTab & Format(Total, "$#0.00")
        Text1.SelStart = Len(Text1.Text) - Len(vbCrLf & "Total:" & vbTab & vbTab & Format(CC, "$#0.00") & vbTab & Format(OB, "$#.00") & vbTab & Format(Total, "$#0.00"))
        Text1.SelLength = Len(vbCrLf & "Total:" & vbTab & vbTab & Format(CC, "$#0.00") & vbTab & Format(OB, "$#0.00") & vbTab & Format(Total, "$#0.00"))
        Text1.SelBold = True
        Text1.SelStart = 0
    End With
    
    Text1.SelStart = 0
    Text1.SelLength = 15
    Text1.SelBold = True
    Text1.SelAlignment = vbCenter
    Text1.SelFontSize = 14
    Text1.SelStart = 16
    Text1.SelLength = 23
    Text1.SelItalic = True
    Text1.SelFontSize = 12
    Text1.SelStart = 42
    Text1.SelLength = 43
    Text1.SelBold = True
    Text1.SelStart = 0
    
    
    Text1.TextRTF = Replace(Text1.TextRTF, "MON-YY", Format(Date, "MMM") & "-" & Format(Date, "yy"))
End Sub

