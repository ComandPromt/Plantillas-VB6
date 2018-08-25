VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   12450
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   240
      Top             =   11820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3075
      Left            =   60
      TabIndex        =   2
      Top             =   8640
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   5424
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Properties"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "SQL Statement"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "rtbSQL"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Database"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frmAccess"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "frmSQL"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.Frame frmSQL 
         Caption         =   "SQL Sever Options"
         Height          =   2475
         Left            =   -72660
         TabIndex        =   25
         Top             =   480
         Width           =   5355
         Begin VB.TextBox txtPWord 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1380
            PasswordChar    =   "*"
            TabIndex        =   33
            Top             =   1755
            Width           =   3255
         End
         Begin VB.TextBox txtUName 
            Height          =   285
            Left            =   1380
            TabIndex        =   32
            Top             =   1275
            Width           =   3255
         End
         Begin VB.TextBox txtDBName 
            Height          =   285
            Left            =   1380
            TabIndex        =   30
            Top             =   795
            Width           =   3255
         End
         Begin VB.TextBox txtSvrname 
            Height          =   285
            Left            =   1380
            TabIndex        =   26
            Top             =   315
            Width           =   3255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Password"
            Height          =   195
            Index           =   3
            Left            =   600
            TabIndex        =   31
            Top             =   1800
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "User Name"
            Height          =   195
            Index           =   2
            Left            =   495
            TabIndex        =   29
            Top             =   1320
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Database Name"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   28
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Server Name"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   27
            Top             =   360
            Width           =   930
         End
      End
      Begin VB.Frame frmAccess 
         Caption         =   "Access Options"
         Height          =   2415
         Left            =   -72600
         TabIndex        =   20
         Top             =   540
         Visible         =   0   'False
         Width           =   5355
         Begin VB.CommandButton cmdGetAccessDB 
            Caption         =   "Find Database"
            Height          =   315
            Left            =   1260
            TabIndex        =   24
            Top             =   1140
            Width           =   2595
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   180
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   660
            Width           =   4935
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Type"
         Height          =   2415
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   2175
         Begin VB.OptionButton optDBType 
            Caption         =   "SQL Sever"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   660
            Width           =   1875
         End
         Begin VB.OptionButton optDBType 
            Caption         =   "MS Access"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   300
            Width           =   1875
         End
      End
      Begin RichTextLib.RichTextBox rtbSQL 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   4
         Top             =   420
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   4471
         _Version        =   393217
         Enabled         =   -1  'True
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form1.frx":0054
      End
      Begin VB.Frame Frame1 
         Height          =   2595
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   7575
         Begin VB.Frame Frame4 
            Caption         =   "Column Properties"
            Height          =   1875
            Left            =   4020
            TabIndex        =   11
            Top             =   300
            Width           =   3375
            Begin VB.Frame Frame5 
               Caption         =   "Font Properties"
               Height          =   1155
               Left            =   1740
               TabIndex        =   16
               Top             =   300
               Width           =   1515
               Begin VB.CheckBox chkFont 
                  Caption         =   "Italic"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   18
                  Top             =   720
                  Width           =   1335
               End
               Begin VB.CheckBox chkFont 
                  Caption         =   "Bold"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   17
                  Top             =   300
                  Width           =   1335
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Text Alignment"
               Height          =   1155
               Left            =   120
               TabIndex        =   12
               Top             =   300
               Width           =   1515
               Begin VB.OptionButton optAlign 
                  Caption         =   "Left"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   15
                  Top             =   240
                  Width           =   915
               End
               Begin VB.OptionButton optAlign 
                  Caption         =   "Center"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   14
                  Top             =   540
                  Width           =   855
               End
               Begin VB.OptionButton optAlign 
                  Caption         =   "Right"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   13
                  Top             =   840
                  Width           =   795
               End
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Report Options"
            Height          =   1995
            Left            =   120
            TabIndex        =   5
            Top             =   180
            Width           =   3795
            Begin VB.TextBox txtTitle 
               Height          =   285
               Left            =   120
               TabIndex        =   8
               Top             =   480
               Width           =   3495
            End
            Begin VB.TextBox txtReportHeading 
               Height          =   315
               Left            =   120
               TabIndex        =   7
               Top             =   1080
               Width           =   3495
            End
            Begin VB.CommandButton cmdPath 
               Caption         =   "Report File Name and Path"
               Height          =   255
               Left            =   720
               TabIndex        =   6
               Top             =   1560
               Width           =   2355
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Report Heading"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   10
               Top             =   840
               Width           =   1125
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Title"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Width           =   300
            End
         End
      End
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Create Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   840
      TabIndex        =   1
      Top             =   11820
      Width           =   6315
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8475
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7755
      ExtentX         =   13679
      ExtentY         =   14949
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsReport As New clsCreateHTMLReport
Private cnMain As New ADODB.Connection

Private Sub chkFont_Click(Index As Integer)
    Select Case Index
        Case 0
            clsReport.Bold = Not clsReport.Bold
        Case 1
            clsReport.Italic = Not clsReport.Italic
    End Select
End Sub

Private Sub cmdGenerate_Click()
    
    Dim strCNN As String
    
    With cnMain
        If .State = 0 Then
            .CommandTimeout = 60
            If optDBType(0).Value = True Then
                .ConnectionString = cdg1.FileName
            End If
            If optDBType(1).Value = True Then
                strCNN = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;"
                strCNN = strCNN & "Initial Catalog=" & txtDBName.Text & ";"
                strCNN = strCNN & "Data Source=" & txtSvrname.Text
                .ConnectionString = strCNN
            End If
            .Open
        End If
    End With

    
    
    
    With clsReport
        Set .ValidConnection = cnMain
        .ReportName = cdg1.FileName
        .SQLStatement = Replace(rtbSQL.Text, vbCrLf, Space(0))
        .Title = txtTitle.Text
        .ReportHeading = txtReportHeading.Text
        If .CreateReport Then
            WebBrowser1.Navigate .ReportName
        Else
            MsgBox "there was an error"
        End If
    
    End With

End Sub

Private Sub cmdGetAccessDB_Click()
    With cdg1
        
        .Filter = "Access (*.mdb)"
        .ShowOpen
    
    End With
    
    
End Sub

Private Sub cmdPath_Click()
    cdg1.Filter = "HTML (*.html)|*.html|HTM (*.htm)|*.htm"
    cdg1.FilterIndex = 1
    cdg1.ShowSave
    
End Sub


Private Sub Form_Resize()

    WebBrowser1.Width = Me.ScaleWidth - 250
    SSTab1.Width = Me.ScaleWidth - 250


End Sub

Private Sub optAlign_Click(Index As Integer)

    Select Case Index
        Case 0
            clsReport.ColumnAlignment = 0
        Case 1
            clsReport.ColumnAlignment = 1
        Case 2
            clsReport.ColumnAlignment = 2
    End Select
End Sub

Private Sub optDBType_Click(Index As Integer)
    Select Case Index
        Case 0
            frmAccess.Top = 480
            frmAccess.Left = 2340
            frmAccess.Visible = True
            frmSQL.Visible = False
        Case 1
            frmAccess.Visible = False
            frmSQL.Top = 480
            frmSQL.Left = 2340
            frmSQL.Visible = True
    End Select
End Sub
