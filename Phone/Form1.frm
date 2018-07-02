VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PhoneBook - Phone Number Database"
   ClientHeight    =   7500
   ClientLeft      =   -1305
   ClientTop       =   720
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   Palette         =   "Form1.frx":030A
   Picture         =   "Form1.frx":2788
   ScaleHeight     =   7500
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   6000
      MaskColor       =   &H00E0E0E0&
      Picture         =   "Form1.frx":271FA
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Sort List Into Alphabetical Order"
      Top             =   0
      Width           =   735
   End
   Begin VB.ListBox lstCountry 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   600
      TabIndex        =   57
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstPostCode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   840
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstState 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   55
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstSuburb 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox lblCountry 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      MousePointer    =   3  'I-Beam
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox lblPostCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6360
      MaxLength       =   4
      MousePointer    =   3  'I-Beam
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox lblState 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      MaxLength       =   3
      MousePointer    =   3  'I-Beam
      TabIndex        =   7
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox lblSuburb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      MousePointer    =   3  'I-Beam
      TabIndex        =   6
      Top             =   1920
      Width           =   4095
   End
   Begin VB.ListBox lstWebSite 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   480
      TabIndex        =   49
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstCoFax 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   480
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstFax 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   47
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstNumbers2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox lblWebSite 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      MousePointer    =   3  'I-Beam
      TabIndex        =   18
      Top             =   5520
      Width           =   3495
   End
   Begin VB.TextBox lblCoFax 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      MousePointer    =   3  'I-Beam
      TabIndex        =   16
      Top             =   5640
      Width           =   4095
   End
   Begin VB.TextBox lblFax 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      MousePointer    =   3  'I-Beam
      TabIndex        =   12
      Top             =   3720
      Width           =   4095
   End
   Begin VB.TextBox lblPhNo2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      MousePointer    =   3  'I-Beam
      TabIndex        =   11
      Top             =   3240
      Width           =   4095
   End
   Begin VB.TextBox lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      MousePointer    =   3  'I-Beam
      TabIndex        =   4
      Top             =   960
      Width           =   4095
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      MaskColor       =   &H00E0E0E0&
      Picture         =   "Form1.frx":2840C
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Delete Entry"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      MaskColor       =   &H00E0E0E0&
      Picture         =   "Form1.frx":2961E
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      MaskColor       =   &H00E0E0E0&
      Picture         =   "Form1.frx":2A830
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Print Entry"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "Form1.frx":2BA42
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Help"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdAddEntry 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Picture         =   "Form1.frx":2CC54
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Add New Entry"
      Top             =   0
      Width           =   735
   End
   Begin VB.ListBox lstNames 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      MaskColor       =   &H00E0E0E0&
      Picture         =   "Form1.frx":2DE66
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Save Data"
      Top             =   0
      Width           =   735
   End
   Begin VB.ListBox lstComments 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstWorkNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   39
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstWork 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstEmail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstMobile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstaddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox lstNumbers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Clear"
      DownPicture     =   "Form1.frx":2F078
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox lblComments 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   6240
      Width           =   8775
   End
   Begin VB.TextBox lblMobile 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      MousePointer    =   3  'I-Beam
      TabIndex        =   13
      Top             =   4200
      Width           =   4095
   End
   Begin VB.TextBox lblWorkNo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      MousePointer    =   3  'I-Beam
      TabIndex        =   15
      Top             =   5160
      Width           =   4095
   End
   Begin VB.TextBox lblAddress 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      MousePointer    =   3  'I-Beam
      TabIndex        =   5
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox lblEmail 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   120
      MousePointer    =   3  'I-Beam
      TabIndex        =   17
      Top             =   4680
      Width           =   3375
   End
   Begin VB.TextBox lblWork 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      MousePointer    =   3  'I-Beam
      TabIndex        =   14
      Top             =   4680
      Width           =   4095
   End
   Begin VB.TextBox lblPhNo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   405
      Left            =   4800
      MousePointer    =   3  'I-Beam
      OLEDropMode     =   1  'Manual
      TabIndex        =   10
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6980
      TabIndex        =   53
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Postcode:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5430
      TabIndex        =   52
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4150
      TabIndex        =   51
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Suburb:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3790
      TabIndex        =   50
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WebSite:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   45
      Top             =   5280
      Width           =   1035
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Co. Fax:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   44
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Fax No.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3710
      TabIndex        =   43
      Top             =   3720
      Width           =   1020
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ph No. 2:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3690
      TabIndex        =   42
      Top             =   3240
      Width           =   1050
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3980
      TabIndex        =   41
      Top             =   960
      Width           =   750
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company No.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3090
      TabIndex        =   34
      Top             =   5160
      Width           =   1650
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3900
      TabIndex        =   33
      Top             =   4200
      Width           =   840
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   31
      Top             =   4440
      Width           =   720
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3520
      TabIndex        =   30
      Top             =   4680
      Width           =   1200
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ph No:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3950
      TabIndex        =   29
      Top             =   2760
      Width           =   795
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3690
      TabIndex        =   28
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name To Search For:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Index           =   1
      NegotiatePosition=   1  'Left
      Begin VB.Menu addentry 
         Caption         =   "&Add Entry"
         Index           =   2
         Shortcut        =   ^A
      End
      Begin VB.Menu SaveData 
         Caption         =   "&Save Data"
         Index           =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu Delete 
         Caption         =   "&Delete Entry"
         Index           =   4
         Shortcut        =   ^D
      End
      Begin VB.Menu sort 
         Caption         =   "Sor&t Entries"
         Shortcut        =   ^T
      End
      Begin VB.Menu NoEntries 
         Caption         =   "N&o. Of Entries"
         Shortcut        =   ^O
      End
      Begin VB.Menu print 
         Caption         =   "&Print"
         Index           =   4
         Shortcut        =   ^P
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
         Index           =   5
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Index           =   6
      NegotiatePosition=   1  'Left
      Begin VB.Menu help2 
         Caption         =   "&Help"
         Index           =   7
         Shortcut        =   ^H
      End
      Begin VB.Menu about 
         Caption         =   "A&bout"
         Index           =   8
         Shortcut        =   ^B
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click(Index As Integer)
    frmSplash.Show
End Sub
Private Sub addentry_Click(Index As Integer)
    frmAddEntry.Show
End Sub
Private Sub cmdAddEntry_Click()
    frmAddEntry.Show
End Sub
Private Sub cmdClear_Click()
    txtSearch.Text = ""
End Sub
Private Sub cmdDelete_Click()
    If lstNames.ListIndex = -1 Then
        If MsgBox("You Do Not Have An Entry Selected", vbExclamation) = vbOK Then Exit Sub
        End If
    If MsgBox("Are You Sure You Want To Delete The Selected Entry?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
    Dim a As Integer
    a = lstNames.ListIndex
    lstNames.RemoveItem a
    lstaddress.RemoveItem a
    lstSuburb.RemoveItem a
    lstState.RemoveItem a
    lstPostCode.RemoveItem a
    lstCountry.RemoveItem a
    lstNumbers.RemoveItem a
    lstNumbers2.RemoveItem a
    lstFax.RemoveItem a
    lstMobile.RemoveItem a
    lstWork.RemoveItem a
    lstWorkNo.RemoveItem a
    lstCoFax.RemoveItem a
    lstEmail.RemoveItem a
    lstWebSite.RemoveItem a
    lstComments.RemoveItem a
    If a = lstNames.ListCount Then
        lstNames.ListIndex = a - 1
    Else
        lstNames.ListIndex = a
    End If
    Call lstNames_Click
    Call Save_It
End Sub
Private Sub cmdExit_Click()
    End
End Sub
Private Sub cmdHelp_Click()
    frmHelp.Show
End Sub
Private Sub cmdPrint_Click()
    If Form1.lstNames.ListIndex = -1 Then
        If MsgBox("Warning: You Do Not Have An Entry Selected. If You Continue, Blank Fields Will Be Printed.", vbCritical) = vbOK Then
        End If
    End If
        If MsgBox("Are You Ready To Print?", vbYesNo) = vbYes Then Call Print_it
End Sub
Private Sub Print_it()
    Printer.Font = "arial"
    Printer.FontSize = 18
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print "PhoneBook"
    Printer.Print ""
    Printer.Font = "Arial"
    Printer.FontUnderline = False
    Printer.FontSize = 12
    Printer.Print "Name: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblName.Text
    Printer.FontBold = True
    Printer.Print "Address: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblAddress.Text
    Printer.FontBold = True
    Printer.Print "Suburb: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblSuburb.Text
    Printer.FontBold = True
    Printer.Print "State: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblState.Text
    Printer.FontBold = True
    Printer.Print "Post Code: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblPostCode.Text
    Printer.FontBold = True
    Printer.Print "Country: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblCountry.Text
    Printer.FontBold = True
    Printer.Print "Phone Number: ",
    Printer.FontBold = False
    Printer.Print Form1.lblPhNo.Text
    Printer.FontBold = True
    Printer.Print "Second Phone Number: ",
    Printer.FontBold = False
    Printer.Print Form1.lblPhNo2.Text
    Printer.FontBold = True
    Printer.Print "Fax Number: ",
    Printer.FontBold = False
    Printer.Print Form1.lblFax.Text
    Printer.FontBold = True
    Printer.Print "Mobile Number: ",
    Printer.FontBold = False
    Printer.Print Form1.lblMobile.Text
    Printer.FontBold = True
    Printer.Print "Company Name: ",
    Printer.FontBold = False
    Printer.Print Form1.lblWork.Text
    Printer.FontBold = True
    Printer.Print "Company Ph. No.: ",
    Printer.FontBold = False
    Printer.Print Form1.lblWorkNo.Text
    Printer.FontBold = True
    Printer.Print "Company Fax Number: ",
    Printer.FontBold = False
    Printer.Print Form1.lblCoFax.Text
    Printer.FontBold = True
    Printer.Print "Email: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblEmail.Text
    Printer.FontBold = True
    Printer.Print "Web Site: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblWebSite.Text
    Printer.FontBold = True
    Printer.Print "Comments: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblComments.Text
    Printer.EndDoc
End Sub
Private Sub cmdSave_Click()
    If MsgBox("Are You Sure You Want to save?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Call Save_It
End Sub
Private Sub Save_It()
    Open "Numbers.dat" For Output As 1
    For i = 0 To lstNames.ListCount - 1
        Print #1, lstNames.List(i)
        Print #1, lstaddress.List(i)
        Print #1, lstSuburb.List(i)
        Print #1, lstState.List(i)
        Print #1, lstPostCode.List(i)
        Print #1, lstCountry.List(i)
        Print #1, lstNumbers.List(i)
        Print #1, lstNumbers2.List(i)
        Print #1, lstFax.List(i)
        Print #1, lstMobile.List(i)
        Print #1, lstWork.List(i)
        Print #1, lstWorkNo.List(i)
        Print #1, lstCoFax.List(i)
        Print #1, lstEmail.List(i)
        Print #1, lstWebSite.List(i)
        Print #1, lstComments.List(i)
    Next i
    Close #1
End Sub
Private Sub Delete_Click(Index As Integer)
    Call cmdDelete_Click
End Sub

Private Sub Exit_Click(Index As Integer)
    End
End Sub
Private Sub SwapList(lst As ListBox, a As Integer, b As Integer)
    Dim temp As String
    temp = lst.List(a)
    lst.List(a) = lst.List(b)
    lst.List(b) = temp
End Sub
Private Sub SwapPeople(a As Integer, b As Integer)
' used by the sort to swap two values
    Call SwapList(lstNames, a, b)
    Call SwapList(lstaddress, a, b)
    Call SwapList(lstSuburb, a, b)
    Call SwapList(lstState, a, b)
    Call SwapList(lstPostCode, a, b)
    Call SwapList(lstCountry, a, b)
    Call SwapList(lstNumbers, a, b)
    Call SwapList(lstNumbers2, a, b)
    Call SwapList(lstFax, a, b)
    Call SwapList(lstMobile, a, b)
    Call SwapList(lstWork, a, b)
    Call SwapList(lstWorkNo, a, b)
    Call SwapList(lstCoFax, a, b)
    Call SwapList(lstEmail, a, b)
    Call SwapList(lstWebSite, a, b)
    Call SwapList(lstComments, a, b)
End Sub
Private Sub Sort_it()
    ' sort them alphabetically
    ' uses a bubble sort
    Dim a As Integer, b As Integer
    For a = 0 To lstNames.ListCount - 2
        For b = a + 1 To lstNames.ListCount - 1
            ' compare and swap if necessary
            If lstNames.List(b) < lstNames.List(a) Then
                Call SwapPeople(a, b)
            End If
        Next b
    Next a
    Call lstNames_Click ' show it all up now again
End Sub
Private Sub cmdSort_Click()
    Call Sort_it
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Dim TempName As String, TempNumber As String
    Open "Numbers.dat" For Input As 1
    On Error Resume Next
    Do Until EOF(1)
        Line Input #1, TempName
        lstNames.AddItem TempName
        Line Input #1, TempAddress
        lstaddress.AddItem TempAddress
        Line Input #1, TempSuburb
        lstSuburb.AddItem TempSuburb
        Line Input #1, tempstate
        lstState.AddItem tempstate
        Line Input #1, TempPostCode
        lstPostCode.AddItem TempPostCode
        Line Input #1, TempCountry
        lstCountry.AddItem TempCountry
        Line Input #1, TempNumber
        lstNumbers.AddItem TempNumber
        Line Input #1, TempNumber2
        lstNumbers2.AddItem TempNumber2
        Line Input #1, TempFax
        lstFax.AddItem TempFax
        Line Input #1, TempMobile
        lstMobile.AddItem TempMobile
        Line Input #1, TempWork
        lstWork.AddItem TempWork
        Line Input #1, TempWorkNo
        lstWorkNo.AddItem TempWorkNo
        Line Input #1, TempCoFax
        lstCoFax.AddItem TempCoFax
        Line Input #1, TempEmail
        lstEmail.AddItem TempEmail
        Line Input #1, TempWebSite
        lstWebSite.AddItem TempWebSite
        Line Input #1, TempComments
        lstComments.AddItem TempComments
    Loop
    Close #1
    lstNames.ListIndex = 0
    Call Sort_it
ErrorHandler:
    Select Case Err.Number
    Case 53
        Call Save_It
    End Select
End Sub
Private Sub help2_Click(Index As Integer)
    frmHelp.Show
End Sub

Private Sub lblComments_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub lblCountry_LostFocus()
If lstNames.ListCount > 0 Then
    lstCountry.List(lstNumbers.ListIndex) = lblCountry.Text
End If
End Sub
Private Sub lblSuburb_LostFocus()
If lstNames.ListCount > 0 Then
    lstSuburb.List(lstNumbers.ListIndex) = lblSuburb.Text
End If
End Sub
Private Sub lblState_LostFocus()
If lstNames.ListCount > 0 Then
    lblState.Text = UCase(lblState.Text)
    lstState.List(lstNumbers.ListIndex) = lblState.Text
End If
End Sub
Private Sub lblPostCode_LostFocus()
If lstNames.ListCount > 0 Then
    lstPostCode.List(lstNumbers.ListIndex) = lblPostCode.Text
End If
End Sub
Private Sub lblCoFax_LostFocus()
If lstNames.ListCount > 0 Then
    lstCoFax.List(lstNumbers.ListIndex) = lblCoFax.Text
End If
End Sub
Private Sub lblFax_LostFocus()
If lstNames.ListCount > 0 Then
    lstFax.List(lstNumbers.ListIndex) = lblFax.Text
End If
End Sub
Private Sub lblMobile_LostFocus()
If lstNames.ListCount > 0 Then
    lstMobile.List(lstNumbers.ListIndex) = lblMobile.Text
End If
End Sub
Private Sub lblEmail_LostFocus()
If lstNames.ListCount > 0 Then
    lstEmail.List(lstNumbers.ListIndex) = lblEmail.Text
End If
End Sub
Private Sub lblName_LostFocus()
If lstNames.ListCount > 0 Then
    lstNames.List(lstNumbers.ListIndex) = lblName.Text
End If
End Sub
Private Sub lblPhNo_LostFocus()
If lstNames.ListCount > 0 Then
   lstNumbers.List(lstNumbers.ListIndex) = lblPhNo.Text
End If
End Sub
Private Sub lblPhNo2_LostFocus()
If lstNames.ListCount > 0 Then
    lstNumbers2.List(lstNumbers.ListIndex) = lblPhNo2.Text
End If
End Sub
Private Sub lblWebSite_LostFocus()
If lstNames.ListCount > 0 Then
   lstWebSite.List(lstNumbers.ListIndex) = lblWebSite.Text
End If
End Sub
Private Sub lblWork_LostFocus()
If lstNames.ListCount > 0 Then
   lstWork.List(lstNumbers.ListIndex) = lblWork.Text
End If
End Sub
Private Sub lblWorkNo_LostFocus()
If lstNames.ListCount > 0 Then
   lstWorkNo.List(lstNumbers.ListIndex) = lblWorkNo.Text
End If
End Sub
Private Sub lblAddress_LostFocus()
If lstNames.ListCount > 0 Then
    lstaddress.List(lstNumbers.ListIndex) = lblAddress.Text
End If
End Sub
Private Sub lblComments_LostFocus()
If lstNames.ListCount > 0 Then
    lstComments.List(lstNumbers.ListIndex) = lblComments.Text
End If
End Sub
Private Sub lstNames_Click()
'On Error GoTo lstNamesErr
   lstNumbers.ListIndex = lstNames.ListIndex
   lstNumbers2.ListIndex = lstNames.ListIndex
   lstFax.ListIndex = lstNames.ListIndex
   lstMobile.ListIndex = lstNames.ListIndex
   lstEmail.ListIndex = lstNames.ListIndex
   lstWork.ListIndex = lstNames.ListIndex
   lstWorkNo.ListIndex = lstNames.ListIndex
   lstCoFax.ListIndex = lstNames.ListIndex
   lstWebSite.ListIndex = lstNames.ListIndex
   lstaddress.ListIndex = lstNames.ListIndex
   lstComments.ListIndex = lstNames.ListIndex
   lstState.ListIndex = lstNames.ListIndex
   lstSuburb.ListIndex = lstNames.ListIndex
   lstCountry.ListIndex = lstNames.ListIndex
   lstPostCode.ListIndex = lstNames.ListIndex
   lblName.Text = lstNames.Text
   lblPhNo.Text = lstNumbers.Text
   lblPhNo2.Text = lstNumbers2.Text
   lblFax.Text = lstFax.Text
   lblMobile.Text = lstMobile.Text
   lblEmail.Text = lstEmail.Text
   lblWork.Text = lstWork.Text
   lblWorkNo.Text = lstWorkNo.Text
   lblCoFax.Text = lstCoFax.Text
   lblWebSite.Text = lstWebSite.Text
   lblAddress.Text = lstaddress.Text
   lblComments.Text = lstComments.Text
   lblSuburb.Text = lstSuburb.Text
   lblState.Text = lstState.Text
   lblCountry.Text = lstCountry.Text
   lblPostCode.Text = lstPostCode.Text
   lblState.Text = UCase(lblState.Text)
'lstNamesErr:
 '   Exit Sub
End Sub
Private Sub NoEntries_Click()
    MsgBox "You Have: " & lstNames.ListCount & " Entries."
End Sub
Private Sub print_Click(Index As Integer)
    Call cmdPrint_Click
End Sub
Private Sub SaveData_Click(Index As Integer)
  If MsgBox("Are You Sure You Want to save?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
  Call Save_It
End Sub
Private Sub sort_Click()
    Call Sort_it
End Sub
Private Sub txtSearch_Change()
    Dim MatchFound As Boolean
    Dim Last As Integer, J As Integer
    lblName.Text = ""
    lblPhNo.Text = ""
    lblPhNo2.Text = ""
    lblFax.Text = ""
    lblMobile.Text = ""
    lblEmail.Text = ""
    lblWork.Text = ""
    lblWorkNo.Text = ""
    lblCoFax.Text = ""
    lblWebSite.Text = ""
    lblAddress.Text = ""
    lblComments.Text = ""
    lblState.Text = ""
    lblPostCode.Text = ""
    lblCountry.Text = ""
    lblSuburb.Text = ""
    Last = lstNames.ListCount - 1
    J = 0
    MatchFound = False
    Do
        If InStr(1, lstNames.List(J), txtSearch.Text, 1) > 0 Then
            MatchFound = True
            lstNames.ListIndex = J
        End If
        J = J + 1
    Loop Until J > Last Or MatchFound
    If Not MatchFound Then
        lstNames.ListIndex = -1
    End If
    
    Call lstNames_Click
End Sub
