VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMain 
   Caption         =   "DataList Demonstration"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      DataField       =   "ISBN"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      DataField       =   "Title"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   2760
      Width           =   2895
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmMain.frx":0000
      DataField       =   "PubID"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Name"
      BoundColumn     =   "PubID"
      Text            =   "DataCombo1"
      Object.DataMember      =   ""
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1080
      Top             =   4200
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=MyBiblio;"
      OLEDBString     =   "DSN=MyBiblio;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Titles"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1080
      Top             =   4560
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   1
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=MyBiblio;"
      OLEDBString     =   "DSN=MyBiblio;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select Name, PubID From Publishers"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "ISBN:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Publisher:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' **************************************
' MS Data List Controls Demonstration
' For Visual Basic Programmer's Journal
' October 1998
' By Jeffrey P. McManus
' jeffreyp@sirius.com
' http://www.redblazer.com/vbdb/
' **************************************
'
'
' Uses MS Data List Controls 6.0.

