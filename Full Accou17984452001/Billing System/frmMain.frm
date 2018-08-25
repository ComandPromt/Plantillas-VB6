VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Billing System   (TO EDIT COMPANY NAME AND DEFAULT INFORMATION, CHANGE THE DefaultInvoice.rtf FILE IN THE SAME DIR AS THE APP)"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtOther 
      Height          =   285
      Left            =   6480
      TabIndex        =   18
      Top             =   5160
      Width           =   3375
   End
   Begin RichTextLib.RichTextBox txtNotes 
      Height          =   1455
      Left            =   4320
      TabIndex        =   63
      Top             =   6720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"FRMMAIN.frx":0000
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Backup Month / Database"
      Height          =   375
      Left            =   7200
      TabIndex        =   62
      Top             =   10440
      Width           =   2775
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Give Current Account New Invoice Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   61
      Top             =   9360
      Width           =   1815
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Done Billing For This Month"
      Height          =   375
      Left            =   4440
      TabIndex        =   60
      Top             =   10440
      Width           =   2775
   End
   Begin VB.PictureBox PicEdited 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1680
      ScaleHeight     =   255
      ScaleWidth      =   975
      TabIndex        =   59
      Top             =   10850
      Width           =   975
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Left            =   6360
      TabIndex        =   57
      Text            =   "Month"
      Top             =   9840
      Width           =   2295
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Delete Last Op."
      Height          =   375
      Left            =   13920
      TabIndex        =   55
      Top             =   10680
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Save Current Operation"
      Height          =   375
      Left            =   12120
      TabIndex        =   53
      Top             =   10680
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Add New Operation"
      Height          =   375
      Left            =   10320
      TabIndex        =   52
      Top             =   10680
      Width           =   1815
   End
   Begin VB.TextBox txtPO 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   10560
      TabIndex        =   43
      Top             =   7080
      Width           =   4575
   End
   Begin VB.TextBox txtAircraft 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   10560
      TabIndex        =   42
      Top             =   6360
      Width           =   4575
   End
   Begin VB.TextBox txtAirbill 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   10560
      TabIndex        =   41
      Top             =   5520
      Width           =   4575
   End
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   10560
      TabIndex        =   40
      Top             =   4800
      Width           =   4575
   End
   Begin VB.TextBox txtDestination 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   10560
      TabIndex        =   39
      Top             =   4080
      Width           =   4575
   End
   Begin VB.TextBox txtWeight 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   10560
      TabIndex        =   38
      Top             =   3360
      Width           =   4575
   End
   Begin VB.TextBox txtOperations 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   10560
      TabIndex        =   37
      Top             =   2520
      Width           =   4575
   End
   Begin VB.TextBox txtFlight 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   10560
      TabIndex        =   36
      Top             =   1800
      Width           =   4575
   End
   Begin VB.TextBox txtDate 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   10560
      TabIndex        =   35
      Top             =   960
      Width           =   4575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "<-- Details"
      Height          =   255
      Left            =   7920
      TabIndex        =   33
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Details -->"
      Height          =   255
      Left            =   8880
      TabIndex        =   32
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Billing Summary"
      Height          =   375
      Left            =   8160
      TabIndex        =   31
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Generate Statement"
      Height          =   375
      Left            =   6240
      TabIndex        =   30
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save Current Account"
      Height          =   375
      Left            =   4320
      TabIndex        =   29
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save New Account"
      Height          =   375
      Left            =   6240
      TabIndex        =   28
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Account"
      Height          =   375
      Left            =   8160
      TabIndex        =   27
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make New Account"
      Height          =   375
      Left            =   4320
      TabIndex        =   26
      Top             =   8400
      Width           =   1815
   End
   Begin VB.ComboBox cboPaid 
      Height          =   315
      ItemData        =   "FRMMAIN.frx":00C9
      Left            =   6480
      List            =   "FRMMAIN.frx":00D3
      TabIndex        =   19
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox txtCityStateZip 
      Height          =   315
      Left            =   6480
      TabIndex        =   17
      Top             =   4680
      Width           =   3375
   End
   Begin VB.TextBox txtBillingAddress 
      Height          =   315
      Left            =   6480
      TabIndex        =   16
      Top             =   4200
      Width           =   3375
   End
   Begin VB.TextBox txtBillingName 
      Height          =   315
      Left            =   6480
      TabIndex        =   15
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Total"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      DataSource      =   "Data1"
      Height          =   315
      Index           =   6
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Rate"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      DataSource      =   "Data1"
      Height          =   315
      Index           =   5
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Outstanding Balance"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      DataSource      =   "Data1"
      Height          =   315
      Index           =   4
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Operations"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Data1"
      Height          =   315
      Index           =   3
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Height          =   315
      Index           =   0
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.ListBox lstAccounts 
      Height          =   10785
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label18 
      Caption         =   "Other Address Info:"
      Height          =   255
      Left            =   4440
      TabIndex        =   64
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label17 
      Caption         =   "Edited This Month:"
      Height          =   255
      Left            =   120
      TabIndex        =   58
      Top             =   10850
      Width           =   1575
   End
   Begin VB.Label Label16 
      Caption         =   "Processing Charges For:"
      Height          =   255
      Left            =   4440
      TabIndex        =   56
      Top             =   9840
      Width           =   1815
   End
   Begin VB.Label lblOP 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10320
      TabIndex        =   54
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label15 
      Caption         =   "Purchase Order Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   51
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label Label14 
      Caption         =   "Aircraft #:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   50
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Airbill #:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   49
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   48
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Destination:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   47
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Weight:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   46
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Operations:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   45
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Flight #:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   44
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   34
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Paid Last Month:"
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   10200
      X2              =   10200
      Y1              =   120
      Y2              =   11040
   End
   Begin VB.Label Label2 
      Caption         =   "Current Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   24
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Billing Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "City/State/Zip:"
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Billing Address:"
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Billing Name:"
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   20
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Total:"
      Height          =   255
      Index           =   6
      Left            =   4440
      TabIndex        =   14
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Rate:"
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   13
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Outstanding Balance:"
      Height          =   255
      Index           =   4
      Left            =   4440
      TabIndex        =   12
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Operations:"
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Invoice #:"
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Current Charges:"
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   9
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Airline:"
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DB As Database
Public Accounts As Recordset
Public AccountInfo As Recordset
Public Options As Recordset
Public Details As Recordset
Public Payments As Recordset
Public Loading As Boolean
Public Op As Integer
Private BackupRS(1 To 2) As Recordset

Private Sub Command1_Click()
    Loading = True
    Me.txtBillingAddress.Text = ""
    Me.txtBillingName.Text = ""
    Me.txtCityStateZip.Text = ""
    Me.txtOther.Text = ""
    Me.txtFields(0).Text = ""
    Me.txtFields(1).Text = "$0.00"
    Me.Options.Edit
    Me.Options.MoveFirst
    Me.txtFields(2).Text = Me.Options!CurrentInvoiceNum
    Me.txtFields(3).Text = "1"
    Me.txtFields(4).Text = "$0.00"
    Me.txtFields(5).Text = "$27.50"
    Me.txtFields(6).Text = "$0.00"
    Me.cboPaid.Text = "Yes"
    lblOP.Caption = ""
    txtFields(0).Locked = False
End Sub

Private Sub Command10_Click()
    SaveDetailData txtFields(0).Text
End Sub

Private Sub Command11_Click()
    Op = Op - 1
    lblOP.Caption = "Current Operation: " & Op
    txtFields(3).Text = Op
    
    ClearDetails
End Sub

Private Sub Command12_Click()
    Dim msgresult As String
    
    msgresult = MsgBox("Are you sure you want to goto the next month?", vbYesNo, "Goto Next Month?")
    If msgresult = vbNo Then Exit Sub
    
    txtPC.Text = MonthName(Format(Date, "m") + 1)
    
    Accounts.MoveFirst
        Do Until Accounts.EOF = True
            Accounts.Edit
            Accounts!Edited = "NO"
            Accounts.Update
            Accounts.MoveNext
        Loop
    Accounts.MoveLast
    Accounts.Edit
    Accounts!Edited = "NO"
    Accounts.Update
'    lstAccounts.ListIndex = 0
'    Do Until lstAccounts.ListIndex = lstAccounts.ListCount + 1
'        Command14_Click  'Backup Account
'        lstAccounts.ListIndex = lstAccounts.ListIndex + 1
'    Loop
End Sub

Private Sub Command13_Click()
    Dim msgresult As String
    
    msgresult = MsgBox("Are you sure you want to give it a new invoice number?", vbYesNo, "Sure?")
    If msgresult = vbNo Then Exit Sub
    
    Options.MoveFirst
    Options.Edit
    txtFields(2).Text = Options!CurrentInvoiceNum
    Options!CurrentInvoiceNum = Int(Options!CurrentInvoiceNum) + 1
    Options.Update
End Sub


Private Sub Command15_Click()
    If lstAccounts.Text = "" Then MsgBox "Please select a company to make a payment for.", , "Select a Company": Exit Sub
    frmPaying.Show vbModal, Me
End Sub

Private Sub Command2_Click()
    Dim Data As String
    
    Data = MsgBox("Are you sure you want to delete " & lstAccounts.Text & "?", vbYesNo, "Delete Account?")
    If Data = vbNo Then Exit Sub
    GotoAirline lstAccounts.Text
    Accounts.Delete
    AccountInfo.Delete
    Details.Delete
    Me.txtBillingAddress.Text = ""
    Me.txtBillingName.Text = ""
    Me.txtCityStateZip.Text = ""
    Me.txtFields(0).Text = ""
    Me.txtFields(1).Text = ""
    Me.txtFields(2).Text = ""
    Me.txtFields(3).Text = ""
    Me.txtFields(4).Text = ""
    Me.txtFields(5).Text = ""
    Me.txtFields(6).Text = ""
    Me.cboPaid.Text = ""
    lstAccounts.Selected(0) = True
    RefreshListbox
End Sub

Private Sub Command3_Click()
    Loading = False
    txtFields(0).Locked = True
    SaveNewData
    RefreshListbox
    lstAccounts.Text = txtFields(0).Text
End Sub

Private Sub Command4_Click()
    SaveData
End Sub

Private Sub Command5_Click()
    frmInvoice.Show vbModal, Me
End Sub

Private Sub Command6_Click()
    frmBS.Show vbModal, Me
End Sub

Public Sub Command7_Click()
    If Op < txtFields(3).Text Then
        Command10_Click 'Save Current One
        Op = Op + 1
        lblOP.Caption = "Current Operation: " & Op
        ClearFields
        LoadDetails txtFields(0).Text
    End If
End Sub

Public Sub Command8_Click()
    If Op > 1 Then
        Command10_Click 'Save Current One
        Op = Op - 1
        lblOP.Caption = "Current Operation: " & Op
        ClearDetails
        LoadDetails txtFields(0).Text
    End If
End Sub

Private Sub Command9_Click()
    Op = Op + 1
    lblOP.Caption = "Current Operation: " & Op
    txtFields(3).Text = Op
    
    ClearDetails
End Sub

Private Sub Form_Load()
    Dim Temp As String
    
    Set DB = OpenDatabase(App.Path & "\DB\AIW-Recordset.mdb")
    Set Accounts = DB.OpenRecordset("Accounts")
    Set AccountInfo = DB.OpenRecordset("Account Information")
    Set Options = DB.OpenRecordset("Options")
    Set Details = DB.OpenRecordset("Details")
    Set Payments = DB.OpenRecordset("Payments")
    Accounts.Edit
    Accounts.MoveFirst
    AccountInfo.Edit
    AccountInfo.MoveFirst
    Options.Edit
    Options.MoveFirst
    Loading = True
    Do Until Accounts.EOF = True
        lstAccounts.AddItem Accounts!Airline
        Accounts.MoveNext
    Loop
    Open App.Path & "\Charges For.txt" For Input As #1
        Line Input #1, Temp
        frmMain.txtPC.Text = Temp
    Close #1
    Loading = False
End Sub

Private Sub Form_LostFocus()
    If txtFields(0).Text <> "" Then Command4_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If txtFields(0).Text <> "" Then Command4_Click
End Sub

Private Sub Label10_Click()
    If txtWeight.Enabled = False Then
        txtWeight.Enabled = True
        txtWeight.BackColor = &H80000005
    Else
        If txtWeight.Text = "" Then
            txtWeight.Enabled = False
            txtWeight.BackColor = &H8000000F
        End If
    End If
End Sub

Private Sub Label11_Click()
    If txtDestination.Enabled = False Then
        txtDestination.Enabled = True
        txtDestination.BackColor = &H80000005
    Else
        If txtDestination.Text = "" Then
            txtDestination.Enabled = False
            txtDestination.BackColor = &H8000000F
        End If
    End If
End Sub

Private Sub Label12_Click()
    If txtDescription.Enabled = False Then
        txtDescription.Enabled = True
        txtDescription.BackColor = &H80000005
    Else
        If txtDescription.Text = "" Then
            txtDescription.Enabled = False
            txtDescription.BackColor = &H8000000F
        End If
    End If
End Sub

Private Sub Label13_Click()
    If txtAirbill.Enabled = False Then
        txtAirbill.Enabled = True
        txtAirbill.BackColor = &H80000005
    Else
        If txtAirbill.Text = "" Then
            txtAirbill.Enabled = False
            txtAirbill.BackColor = &H8000000F
        End If
    End If
End Sub

Private Sub Label14_Click()
    If txtAircraft.Enabled = False Then
        txtAircraft.Enabled = True
        txtAircraft.BackColor = &H80000005
    Else
        If txtAircraft.Text = "" Then
            txtAircraft.Enabled = False
            txtAircraft.BackColor = &H8000000F
        End If
    End If
End Sub

Private Sub Label15_Click()
    If txtPO.Enabled = False Then
        txtPO.Enabled = True
        txtPO.BackColor = &H80000005
    Else
        If txtPO.Text = "" Then
            txtPO.Enabled = False
            txtPO.BackColor = &H8000000F
        End If
    End If
End Sub

Private Sub Label6_Click()
    If txtDate.Enabled = False Then
        txtDate.Enabled = True
        txtDate.BackColor = &H80000005
        txtDate.Text = (Format(Date, "mm") - 1) & Format(Date, "/dd/yy")
        txtDate.SelStart = 2
        txtDate.SelLength = 2
        txtDate.SetFocus
    Else
        If txtDate.Text = "" Then
            txtDate.Enabled = False
            txtDate.BackColor = &H8000000F
        End If
    End If
End Sub

Private Sub Label8_Click()
    If txtFlight.Enabled = False Then
        txtFlight.Enabled = True
        txtFlight.BackColor = &H80000005
    Else
        If txtFlight.Text = "" Then
            txtFlight.Enabled = False
            txtFlight.BackColor = &H8000000F
        End If
    End If
End Sub

Private Sub Label9_Click()
    If txtOperations.Enabled = False Then
        txtOperations.Enabled = True
        txtOperations.BackColor = &H80000005
    Else
        If txtOperations.Text = "" Then
            txtOperations.Enabled = False
            txtOperations.BackColor = &H8000000F
        End If
    End If
End Sub

Private Sub lstAccounts_Click()
    
    Dim SaveName As String
    Dim SaveIndex As Integer
    Dim Edited As String
    
    SaveName = lstAccounts.Text
    SaveIndex = lstAccounts.ListIndex

    If txtFields(0).Text <> "" Then
        Command4_Click
    End If
    
    LoadData SaveName
    Op = 1
    LoadDetails txtFields(0).Text
    If txtFields(3).Text > 0 Then
        Op = 1
    Else
        Op = 0
    End If
    Me.Refresh
    Open App.Path & "\Charges For.txt" For Output As #1
        Print #1, frmMain.txtPC.Text
    Close #1
    
    On Error Resume Next
    Edited = Accounts!Edited
    If UCase(Edited) = "YES" Then
        PicEdited.BackColor = vbGreen
    Else
        PicEdited.BackColor = vbRed
    End If
End Sub

Private Sub PicEdited_Click()
    If GotoAirline(lstAccounts.Text) = False Then Exit Sub
    Accounts.Edit
    If PicEdited.BackColor = vbRed Then
        PicEdited.BackColor = vbGreen
        Accounts("Edited") = "YES"
    Else
        PicEdited.BackColor = vbRed
        Accounts("Edited") = "NO"
    End If
    Accounts.Update
End Sub

Private Sub txtBillingName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then txtBillingName.Text = txtFields(0).Text
End Sub

Private Sub txtFields_Change(Index As Integer)
    If Loading = False And Index = 3 Then
        SaveData
    End If
'    If IsNumeric(txtFields(3).Text) = True And IsNumeric(txtFields(5).Text) = True And IsNumeric(txtFields(4).Text) = True And IsNumeric(txtFields(1).Text) = True Then
'        Dim ab, bc As Double
'        ab = Round((CDbl(txtFields(3).Text) * CDbl(txtFields(5).Text)), 2)
'        bc = Round(CDbl(txtFields(4).Text) + CDbl(txtFields(1).Text), 2)
'        txtFields(1).Text = Format(ab, "$#0.00")
'        txtFields(6).Text = Format(bc, "$#0.00")
'    End If
End Sub

Private Sub txtFields_Click(Index As Integer)
    Dim Temp As String
    
    If Loading = True Then Exit Sub
    Select Case Index
    Case 4:
        Temp = InputBox("Please enter a new outstanding balance:", "New Outstanding Balance", txtFields(4).Text)
        If Temp = "" Or Temp = txtFields(4).Text Then Exit Sub
        txtFields(4).Text = Temp
    Case 5:
        Temp = InputBox("Please type in a new rate:", "Rate", txtFields(5).Text)
        If Temp = "" Or Temp = txtFields(5).Text Then Exit Sub
        txtFields(5).Text = Temp
    End Select
End Sub

