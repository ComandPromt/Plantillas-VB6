VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fTestTree 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Test Class Wrapper - cTreeView : ADO CODE EXAMPLE"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fTestTree.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAction 
      Caption         =   "Action:"
      Height          =   1065
      Left            =   3465
      TabIndex        =   9
      Top             =   105
      Width           =   3270
      Begin VB.CommandButton cmdDialog 
         Caption         =   "Add"
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   10
         Top             =   315
         Width           =   960
      End
      Begin VB.CommandButton cmdDialog 
         Caption         =   "Rename"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   11
         Top             =   630
         Width           =   960
      End
      Begin VB.CommandButton cmdDialog 
         Caption         =   "Move"
         Height          =   315
         Index           =   2
         Left            =   1155
         TabIndex        =   12
         Top             =   315
         Width           =   960
      End
      Begin VB.CommandButton cmdDialog 
         Caption         =   "Copy"
         Height          =   315
         Index           =   3
         Left            =   1155
         TabIndex        =   13
         Top             =   630
         Width           =   960
      End
      Begin VB.CommandButton cmdDialog 
         Caption         =   "Delete"
         Height          =   315
         Index           =   4
         Left            =   2205
         TabIndex        =   14
         Top             =   315
         Width           =   960
      End
   End
   Begin MSComctlLib.ImageList imgDialog 
      Left            =   2625
      Top             =   1890
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTestTree.frx":1782
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTestTree.frx":1D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTestTree.frx":22B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTestTree.frx":2850
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTestTree.frx":2DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTestTree.frx":3384
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTestTree.frx":369E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTestTree.frx":39B8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkDialog 
      Caption         =   "Hot Tracking Cursor"
      Height          =   195
      Index           =   2
      Left            =   7455
      TabIndex        =   64
      Top             =   6570
      Value           =   1  'Checked
      Width           =   1800
   End
   Begin VB.CheckBox chkDialog 
      Caption         =   "Allow Label Edit"
      Height          =   195
      Index           =   1
      Left            =   5512
      TabIndex        =   63
      Top             =   6570
      Value           =   1  'Checked
      Width           =   1485
   End
   Begin VB.CheckBox chkDialog 
      Caption         =   "Allow Drag'n'Drop"
      Height          =   195
      Index           =   0
      Left            =   3465
      TabIndex        =   62
      Top             =   6570
      Value           =   1  'Checked
      Width           =   1590
   End
   Begin VB.Frame fraFind 
      Caption         =   "Search:"
      Height          =   1380
      Left            =   105
      TabIndex        =   1
      Top             =   0
      Width           =   3165
      Begin VB.CommandButton cmdFind 
         Caption         =   "Prev"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   2310
         TabIndex        =   8
         Top             =   945
         Width           =   750
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Next"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1207
         TabIndex        =   7
         Top             =   945
         Width           =   750
      End
      Begin VB.OptionButton optFind 
         Caption         =   "Product"
         Height          =   315
         Index           =   1
         Left            =   1890
         TabIndex        =   5
         Top             =   597
         Width           =   855
      End
      Begin VB.OptionButton optFind 
         Caption         =   "Group"
         Height          =   315
         Index           =   0
         Left            =   735
         TabIndex        =   4
         Top             =   597
         Width           =   855
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "First"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   6
         Top             =   945
         Width           =   750
      End
      Begin VB.TextBox txtFind 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   735
         TabIndex        =   3
         Top             =   250
         Width           =   2325
      End
      Begin VB.Label lblFind 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Find: "
         Height          =   315
         Left            =   105
         TabIndex        =   2
         Top             =   250
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "End"
      Height          =   315
      Index           =   5
      Left            =   2745
      TabIndex        =   61
      Top             =   6510
      Width           =   527
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "PDn"
      Height          =   315
      Index           =   4
      Left            =   2220
      TabIndex        =   60
      Top             =   6510
      Width           =   527
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Dn"
      Height          =   315
      Index           =   3
      Left            =   1680
      TabIndex        =   59
      Top             =   6510
      Width           =   527
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Up"
      Height          =   315
      Index           =   2
      Left            =   1155
      TabIndex        =   58
      Top             =   6510
      Width           =   527
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "PUp"
      Height          =   315
      Index           =   1
      Left            =   630
      TabIndex        =   57
      Top             =   6510
      Width           =   527
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Hm"
      Height          =   315
      Index           =   0
      Left            =   105
      TabIndex        =   56
      Top             =   6510
      Width           =   527
   End
   Begin MSComctlLib.TreeView tvwDialog 
      Height          =   4950
      Left            =   105
      TabIndex        =   0
      Top             =   1470
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   8731
      _Version        =   393217
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraDetail 
      Caption         =   "Selected Node Details:"
      Height          =   3585
      Left            =   3465
      TabIndex        =   25
      Top             =   2835
      Width           =   5790
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   14
         Left            =   4305
         TabIndex        =   55
         Top             =   3150
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   13
         Left            =   4305
         TabIndex        =   53
         Top             =   2835
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   12
         Left            =   4305
         TabIndex        =   51
         Top             =   2520
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   11
         Left            =   4305
         TabIndex        =   49
         Top             =   2205
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   10
         Left            =   4305
         TabIndex        =   47
         Top             =   1890
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   9
         Left            =   4305
         TabIndex        =   45
         Top             =   1575
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   8
         Left            =   4305
         TabIndex        =   43
         Top             =   1260
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Index           =   7
         Left            =   1470
         TabIndex        =   29
         Top             =   630
         Width           =   4215
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   6
         Left            =   1470
         TabIndex        =   41
         Top             =   2835
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   5
         Left            =   1470
         TabIndex        =   39
         Top             =   2520
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   4
         Left            =   1470
         TabIndex        =   37
         Top             =   2205
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   3
         Left            =   1470
         TabIndex        =   35
         Top             =   1890
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   2
         Left            =   1470
         TabIndex        =   33
         Top             =   1575
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   1470
         TabIndex        =   31
         Top             =   1260
         Width           =   1380
      End
      Begin VB.Label lblDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1470
         TabIndex        =   27
         Top             =   315
         Width           =   4215
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Is Visible: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   14
         Left            =   2940
         TabIndex        =   54
         Top             =   3150
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Sorted: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   13
         Left            =   2940
         TabIndex        =   52
         Top             =   2835
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Selected: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   12
         Left            =   2940
         TabIndex        =   50
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Root Node:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   11
         Left            =   2940
         TabIndex        =   48
         Top             =   2205
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Parent: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   10
         Left            =   2940
         TabIndex        =   46
         Top             =   1890
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Key: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   9
         Left            =   2940
         TabIndex        =   44
         Top             =   1575
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Index: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   8
         Left            =   2940
         TabIndex        =   42
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Full Path: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   7
         Left            =   105
         TabIndex        =   28
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Fore Colour: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   6
         Left            =   105
         TabIndex        =   40
         Top             =   2835
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Expanded: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   5
         Left            =   105
         TabIndex        =   38
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Has Children: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   4
         Left            =   105
         TabIndex        =   36
         Top             =   2205
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Is Checked: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   3
         Left            =   105
         TabIndex        =   34
         Top             =   1890
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Is Bold: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   2
         Left            =   105
         TabIndex        =   32
         Top             =   1575
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Back Colour: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   1
         Left            =   105
         TabIndex        =   30
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Text: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Index           =   0
         Left            =   105
         TabIndex        =   26
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.Frame fraDialog 
      Height          =   1485
      Left            =   3465
      TabIndex        =   17
      Top             =   1230
      Width           =   5790
      Begin VB.CheckBox chkDialog 
         Caption         =   "Is Product?"
         Height          =   225
         Index           =   3
         Left            =   4515
         TabIndex        =   20
         ToolTipText     =   "Is the new node a Group or a Product?"
         Top             =   255
         Visible         =   0   'False
         Width           =   1190
      End
      Begin VB.TextBox txtDialog 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1575
         TabIndex        =   22
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.CommandButton cmdDialog 
         Caption         =   "Go"
         Height          =   330
         Index           =   5
         Left            =   4515
         TabIndex        =   23
         Top             =   622
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtDialog 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1575
         TabIndex        =   19
         Top             =   210
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.Label lblComments 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   210
         TabIndex        =   24
         Top             =   1050
         Width           =   5370
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Node Text: "
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   18
         Top             =   210
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Parent Node: "
         Height          =   315
         Index           =   1
         Left            =   210
         TabIndex        =   21
         Top             =   630
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Label lblEvent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000010&
      Caption         =   "*Action* Time:"
      ForeColor       =   &H80000014&
      Height          =   300
      Index           =   2
      Left            =   6825
      TabIndex        =   66
      Top             =   735
      Width           =   1170
   End
   Begin VB.Label lblEvent 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   8085
      TabIndex        =   65
      Top             =   735
      Width           =   1170
   End
   Begin VB.Label lblEvent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   8085
      TabIndex        =   16
      Top             =   315
      Width           =   1170
   End
   Begin VB.Label lblEvent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000010&
      Caption         =   "Last Event: "
      ForeColor       =   &H80000014&
      Height          =   300
      Index           =   0
      Left            =   6825
      TabIndex        =   15
      Top             =   315
      Width           =   1170
   End
   Begin VB.Menu mnuPopNode 
      Caption         =   "PopNode"
      Visible         =   0   'False
      Begin VB.Menu mnuNode 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu mnuNode 
         Caption         =   "&Move"
         Index           =   1
      End
      Begin VB.Menu mnuNode 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuNode 
         Caption         =   "C&ut"
         Index           =   3
      End
      Begin VB.Menu mnuNode 
         Caption         =   "&Copy"
         Index           =   4
      End
      Begin VB.Menu mnuNode 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnuNode 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuNode 
         Caption         =   "&Delete"
         Index           =   7
      End
      Begin VB.Menu mnuNode 
         Caption         =   "&Rename"
         Index           =   8
      End
      Begin VB.Menu mnuNode 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuPopTree 
         Caption         =   "&Treeview"
         Begin VB.Menu mnuTree 
            Caption         =   "&Home"
            Index           =   0
         End
         Begin VB.Menu mnuTree 
            Caption         =   "&Page Up"
            Index           =   1
         End
         Begin VB.Menu mnuTree 
            Caption         =   "&Up"
            Index           =   2
         End
         Begin VB.Menu mnuTree 
            Caption         =   "&Down"
            Index           =   3
         End
         Begin VB.Menu mnuTree 
            Caption         =   "P&age Down"
            Index           =   4
         End
         Begin VB.Menu mnuTree 
            Caption         =   "&End"
            Index           =   5
         End
         Begin VB.Menu mnuTree 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuTree 
            Caption         =   "All&ow Drag'n'Drop"
            Checked         =   -1  'True
            Index           =   7
         End
         Begin VB.Menu mnuTree 
            Caption         =   "Allow &Label Edit"
            Checked         =   -1  'True
            Index           =   8
         End
         Begin VB.Menu mnuTree 
            Caption         =   "&Hot Tracking Cursor"
            Checked         =   -1  'True
            Index           =   9
         End
      End
   End
End
Attribute VB_Name = "fTestTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fTestTree
' Author:       Graeme Grant
' Date:         28/09/2001
' Version:      01.00.03            (Version is inline with cTREEVIEW class)
' Description:  Test/Demo TreeView Handler
' Edit History: 01.00.00 18/09/2001 Initial Release
'               01.00.01 21/09/2001 Added Pop-up Menus
'               01.00.01 21/09/2001 Added Icons
'               01.00.01 21/09/2001 Fixed bug with find code not enabling
'                                   next/previous buttons and continuing the
'                                   search under certain conditions.
'               01.00.03 28/09/2001 Added benchmarking for testing purposes
'                                   to maximise performance
'               01.00.03 28/09/2001 Add new 'Copy Node' Feature
'               01.00.03 28/09/2001 Added pAddNode & pAddRecord
'               01.00.03 28/09/0201 Added Action timer code to track
'                                   performance of routines
'               01.00.03 28/09/2001 Added Cut, Copy, Paste to Node Menu
'               01.00.03 28/09/2001 Made Treeview context menu a submenu of
'                                   node contect menu
'               01.00.03 28/09/2001 Adjusted pExpand Node to handle group
'                                   nodes without 'Dummy' hidden nodes
'
' Notes:        01.00.01 21/09/2001 When working with TreeView & ImageList
'                                   controls with icons, the ImageList
'                                   control doesn't check the screen colour
'                                   depth - it just grabs the first icon
'                                   of the designated size from the captured
'                                   icon file. Not the displayed icon when
'                                   viewing the contents of the ImageList
'                                   control. Threrefore to have the correct
'                                   icon displayed in the TreeView control
'                                   you must extract the correct image from
'                                   the icon (.ICO) file using an icon
'                                   extraction utility.
'
'===========================================================================

Option Explicit

'===========================================================================
' Private: Variables and Declarations
'
Private Const clPRODCOLOR As Long = &H800080    '@@ v01.00.03

Private WithEvents moTree As cTreeView
Attribute moTree.VB_VarHelpID = -1
Private moSelectedNode    As MSComctlLib.Node
Private moDestNode        As MSComctlLib.Node
Private msDragTarget      As String
Private msNodeText        As String
Private meMode            As eCommand
Private meFocus           As eTextBox
Private mbIsDirty         As Boolean
Private mbCutOperation    As Boolean            '@@ v01.00.03
Private msCopyKey         As String             '@@

Private Enum eNodeMenu                          '@@ v01.00.03
    [Menu Add] = 0
    [Menu Move] = 1
    [Menu Cut] = 3
    [Menu Copy] = 4
    [Menu Paste] = 5
    [Menu Delete] = 7
    [Menu Rename] = 8
End Enum

Private Enum eCommand
    [Add Node] = 0
    [Rename Node] = 1
    [Move Node] = 2
    [Copy Node] = 3
    [Delete Node] = 4
    [Execute Mode] = 5
End Enum

Private Enum eTextBox
    [No Selection] = -1
    [Node Text] = 0
    [Parent Node] = 1
End Enum

Private Enum eCheck
    [Drag Drop] = 0
    [Label Edit] = 1
    [HotTracking] = 2
    [Action Option] = 3                         '@@ v01.00.03
End Enum

Private Enum eCommandFind
    [Find First] = 0
    [Find Next] = 1
    [Find Previous] = 2
End Enum

Private Enum eFindMode
    [Group] = 0
    [Product] = 1
End Enum

'===========================================================================
' Private: ADO Declarations
'
'## Get Groups by Group ID
Private Const mcSQL_GRP1  As String = "SELECT DISTINCTROW Desc, GroupID, PkID " + _
                                      "FROM [Group] " + _
                                      "WHERE ((Active)=True) and ((Type)=0) and ((GroupID)="
Private Const mcSQL_GRP2  As String = ") ORDER BY GroupID, SeqNum, PkID"

'## Get Products by Group ID
Private Const mcSQL_PROD1 As String = "SELECT DISTINCTROW Desc, PkID " + _
                                      "FROM [Product] " + _
                                      "WHERE ((GroupID)="
Private Const mcSQL_PROD2 As String = ") and ((Active)=True) " + _
                                      "ORDER BY Code"

'## Add Group                                   '@@ v01.00.03
Private Const mcSQL_AGRP  As String = "INSERT INTO [GROUP] ([Desc], [GroupID], [Type], [Active])" + _
                                      "VALUES (?, ?, ?, ?) "
'## Update Group Link ID
Private Const mcSQL_UGRP1 As String = "UPDATE DISTINCTROW [Group] " + _
                                      "SET [Group].GroupID = "
Private Const mcSQL_UGRP2 As String = " WHERE (((Group.PkID)="
Private Const mcSQL_UGRP3 As String = "))"

'## Rename Group
Private Const mcSQL_RGRP1 As String = "UPDATE DISTINCTROW [Group] " + _
                                      "SET [Group].Desc = '"
Private Const mcSQL_RGRP2 As String = "' WHERE (((Group.PkID)="
Private Const mcSQL_RGRP3 As String = "))"

'## Delete Group
Private Const mcSQL_DGRP1 As String = "DELETE DISTINCTROW Group.PkID " + _
                                      "FROM [Group] " + _
                                      "WHERE (((Group.PkID)="
Private Const mcSQL_DGRP2 As String = "))"

'## Update Product Group Link ID
Private Const mcSQL_UPRD1 As String = "UPDATE DISTINCTROW Product " + _
                                      "SET Product.GroupID = "
Private Const mcSQL_UPRD2 As String = " WHERE (((Product.PkID)="
Private Const mcSQL_UPRD3 As String = "))"

'## Add Product                                 '@@ v01.00.03
Private Const mcSQL_APRD  As String = "INSERT INTO [Product] ([Desc], [GroupID], [Active])" + _
                                      "VALUES (?, ?, ?) "

'## Rename Product
Private Const mcSQL_RPRD1 As String = "UPDATE DISTINCTROW Product " + _
                                      "SET Product.Desc = '"
Private Const mcSQL_RPRD2 As String = "' WHERE (((Product.PkID)="
Private Const mcSQL_RPRD3 As String = "))"

'## Delete Product
Private Const mcSQL_DPRD1 As String = "DELETE DISTINCTROW Product.PkID " + _
                                      "FROM Product " + _
                                      "WHERE (((Product.PkID)="
Private Const mcSQL_DPRD2 As String = "))"

'## Find Group By Desc
Private Const mcSQL_FGRP  As String = "SELECT DISTINCTROW PkID, Desc, GroupID " + _
                                      "FROM [Group] " + _
                                      "WHERE ((Active)=True)  AND (Type=0) " + _
                                      "ORDER BY GroupID, SeqNum, PkID"

'## Find Product By Desc
Private Const mcSQL_FPRD  As String = "SELECT DISTINCTROW PkID, Desc, GroupID " + _
                                      "FROM [Product] " + _
                                      "WHERE ((Active)=True) " + _
                                      "ORDER BY Code"

Private moFindRS    As ADODB.Recordset
Private moGrpRS     As ADODB.Recordset
Private mbFindNext  As Boolean
Private meFindMode  As eFindMode
Private mcDB        As cDB
Private mbIsBusy    As Boolean

'===========================================================================
' Form Events
'
Private Sub chkDialog_Click(Index As Integer)

    Dim bState As Boolean

    bState = CBool(chkDialog(Index).Value = 1)
    With moTree
        Select Case Index                                               '@@ v01.00.03
            Case [Action Option]                                        '@@
                '## Just in case we want to do something here....
            Case Else                                                   '@@
                Select Case Index
                    Case [Drag Drop]:   .DragEnabled = bState
                    Case [Label Edit]:  .Ctrl.LabelEdit = CByte(Abs(bState = False))
                    Case [HotTracking]: .Ctrl.HotTracking = bState
                End Select
                '
                '## mnuTree Pop-up Menu enable/disable checks
                '
                mnuTree(Index + 7).Checked = chkDialog(Index).Value     '@@ v01.00.01 Adjust pop menu checks
            End Select                                                  '@@ v01.00.03
    End With
End Sub

Private Sub cmdDialog_Click(Index As Integer)
    '
    '## Add/Rename/Move/Copy/Delete/Execute
    '
    Action Index

End Sub

Private Sub cmdDialog_GotFocus(Index As Integer)
    pShowEvent ""
    meFocus = [No Selection]
End Sub

Private Sub cmdMove_Click(Index As Integer)
    
    Dim oTmr As cBenchmark                                  '@@ v01.00.03
    Set oTmr = New cBenchmark                               '@@
    oTmr.Start                                              '@@
    
    With moTree
        .ScrollView CByte(Index)                            '@@ cTREEVIEW.CLS (v01.00.01) example
        .Ctrl.SetFocus
    End With

    oTmr.Finish                                             '@@ v01.00.03
    pShowTimer oTmr.ElapsedTime                             '@@
    Select Case Index                                       '@@
        Case [Home]:      pShowEvent "*Move: Home*"         '@@
        Case [Page Up]:   pShowEvent "*Move: PgUp*"         '@@
        Case [Up]:        pShowEvent "*Move: Up*"           '@@
        Case [Down]:      pShowEvent "*Move: Down*"         '@@
        Case [Page Down]: pShowEvent "*Move: PgDn*"         '@@
        Case [End]:       pShowEvent "*Move: End*"          '@@
    End Select                                              '@@

End Sub

Private Sub Form_Load()

    Set moTree = New cTreeView

    With moTree
        '
        '## Hook treeview control
        '
        .HookCtrl tvwDialog
        .Redraw False
        '
        '## Set TreeView features
        '
        With .Ctrl
            .Style = tvwTreelinesPlusMinusPictureText
            .LineStyle = tvwRootLines
            .Indentation = 10
            .ImageList = imgDialog
            .FullRowSelect = False
            .HideSelection = False
            .HotTracking = True
            '##--- ADO Code Start ----------------------------------
            '
            '## Build TreeView data
            '
            pInitData
            '
            '##--- ADO Code End ------------------------------------
            pShowNodeDetails .SelectedItem
        End With
        .ContextMenuMode = [After Click]
        .DragEnabled = True
        .Redraw True
    End With
    meFocus = [No Selection]        '## Set Textbox focus to nothing
    optFind([Product]).Value = True '## Set find option default to product
    FlatBorder txtFind.hwnd         '## Give the find box a flat border

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moTree = Nothing
End Sub

Private Sub mnuNode_Click(Index As Integer)                     '@@ v01.00.01

    'Set moSelectedNode = tvwDialog.SelectedItem                 '## Remember selected node
    Select Case Index
        Case [Menu Add]
            moTree.CutIconState False                           '@@ v01.00.03
            Set moSelectedNode = tvwDialog.SelectedItem         '## Remember selected node
            cmdDialog_Click [Add Node]                          '## Press Add button
            txtDialog([Parent Node]).Text = moSelectedNode.Text '## Update screen with selected item
            txtDialog([Node Text]).SetFocus                     '## Set focus to the next step
            mbCutOperation = False                              '@@ v01.00.03
            mnuNode([Menu Paste]).Enabled = False               '@@

        Case [Menu Rename]
            moTree.CutIconState False                           '@@ v01.00.03
            Set moSelectedNode = tvwDialog.SelectedItem         '## Remember selected node
            cmdDialog_Click [Rename Node]                       '## Press Rename button
            txtDialog([Node Text]).Text = moSelectedNode.Text   '## Update screen with selected item
            txtDialog([Parent Node]).SetFocus                   '## Set focus to the next step
            mbCutOperation = False                              '@@ v01.00.03
            mnuNode([Menu Paste]).Enabled = False               '@@

        Case [Menu Move]
            moTree.CutIconState False                           '@@ v01.00.03
            Set moSelectedNode = tvwDialog.SelectedItem         '## Remember selected node
            cmdDialog_Click [Move Node]                         '## Press Move button
            txtDialog([Node Text]).Text = moSelectedNode.Text   '## Update screen with selected item
            txtDialog([Parent Node]).SetFocus                   '## Set focus to the next step
            mbCutOperation = False                              '@@ v01.00.03
            mnuNode([Menu Paste]).Enabled = False               '@@

        Case [Menu Cut]                                         '@@ v01.00.03
            moTree.CutIconState True                            '@@
            Set moSelectedNode = tvwDialog.SelectedItem         '@@ Remember selected node
            mbCutOperation = True                               '@@
            cmdDialog_Click [Move Node]                         '@@ Press Move button
            txtDialog([Node Text]).Text = moSelectedNode.Text   '@@ Update screen with selected item
            txtDialog([Parent Node]).SetFocus                   '@@ Set focus to the next step
            mnuNode([Menu Paste]).Enabled = True                '@@

        Case [Menu Paste]                                       '@@ v01.00.03
            moTree.CutIconState False                           '@@
            Set moDestNode = tvwDialog.SelectedItem             '@@ Remember selected node
            txtDialog([Parent Node]).Text = moDestNode.Text     '@@
            If meMode = [Copy Node] Then
                chkDialog([Action Option]).Value = _
                    Abs(CInt(MsgBox("Include all child nodes in Copy Paste operation?", _
                                    vbDefaultButton1 + vbQuestion + vbYesNo, _
                                    "Copy Node(s)") = vbYes))   '@@
            End If                                              '@@
            cmdDialog_Click [Execute Mode]                      '@@ Execute action
            mnuNode([Menu Paste]).Enabled = False               '@@

        Case [Menu Copy]                                        '@@ v01.00.03
            moTree.CutIconState False                           '@@
            Set moSelectedNode = tvwDialog.SelectedItem         '@@ Remember selected node
            cmdDialog_Click [Copy Node]                         '@@ Press Copy button
            txtDialog([Node Text]).Text = moSelectedNode.Text   '@@ Update screen with selected item
            txtDialog([Parent Node]).SetFocus                   '@@ Set focus to the next step
            mbCutOperation = False                              '@@
            mnuNode([Menu Paste]).Enabled = True                '@@

        Case [Menu Delete]
            moTree.CutIconState False                           '@@ v01.00.03
            Set moSelectedNode = tvwDialog.SelectedItem         '## Remember selected node
            cmdDialog_Click [Delete Node]                       '## Press Delete button
            txtDialog([Node Text]).Text = moSelectedNode.Text   '## Update screen with selected item
            cmdDialog_Click [Execute Mode]                      '## Press Go button & start delete action
            mbCutOperation = False                              '@@ v01.00.03
            mnuNode([Menu Paste]).Enabled = False               '@@

    End Select

End Sub

Private Sub mnuTree_Click(Index As Integer)     '@@ v01.00.01

    Select Case Index
        Case 0 To 5                             '## Home/PgUp/Up/ Down/PgDn/End
            cmdMove_Click Index                 '## Press designated TreeView Scroll Button
        Case 7 To 9                             '## Enable/Disable Drag'n'Drop/Label Edit/Hot Tracking
            With chkDialog(Index - 7)
                .Value = Abs(Not CBool(.Value)) '## Flip value
            End With
            chkDialog_Click Index - 7           '## Raise chkDialog Event
    End Select

End Sub

Private Sub tvwDialog_AfterLabelEdit(Cancel As Integer, NewString As String)
    '
    '## Note: TreeView events can also be directly handled from the form
    '
    '##--- ADO Code Start ----------------------------------
    If NewString <> msNodeText Then
        If Not pRenameRecord(tvwDialog.SelectedItem, NewString) Then
            MsgBox "Unable to rename the selected node.", _
                   vbApplicationModal + vbExclamation + vbOKOnly, _
                   "Rename Node"
            moSelectedNode.Text = msNodeText
        End If
        msNodeText = ""
    End If
    '##--- ADO Code End ------------------------------------
    pShowEvent "*Rename*"

End Sub

Private Sub tvwDialog_BeforeLabelEdit(Cancel As Integer)
    Set moSelectedNode = tvwDialog.SelectedItem
    msNodeText = moSelectedNode.Text
    pShowEvent "*Edit*"
End Sub

Private Sub tvwDialog_Collapse(ByVal Node As MSComctlLib.Node)
    Dim oTmr As cBenchmark                                  '@@ v01.00.03
    Set oTmr = New cBenchmark                               '@@
    oTmr.Start                                              '@@
    oTmr.Finish                                             '@@
    pShowTimer oTmr.ElapsedTime                             '@@
    pShowNodeDetails Node
    pShowEvent "*Collapsed*"
End Sub

Private Sub tvwDialog_Expand(ByVal Node As MSComctlLib.Node)

    Dim oTmr As cBenchmark                                  '@@ v01.00.03
    Set oTmr = New cBenchmark                               '@@

    pShowNodeDetails Node
    pShowEvent "*Expanded*"

    oTmr.Start                                              '@@ v01.00.03
    '##--- ADO Code Start ----------------------------------
    pExpandNode Node
    '##--- ADO Code End ------------------------------------
    oTmr.Finish                                             '@@ v01.00.03
    pShowTimer oTmr.ElapsedTime                             '@@
End Sub

Private Sub tvwDialog_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete                                        '## Delete Node
            Set moSelectedNode = tvwDialog.SelectedItem         '## Remember selected node
            cmdDialog_Click [Delete Node]                       '## Press Delete button
            txtDialog([Node Text]).Text = moSelectedNode.Text   '## Update screen with selected item
            cmdDialog_Click [Execute Mode]                      '## Press Go button & start delete action
    End Select
End Sub

Private Sub txtDialog_GotFocus(Index As Integer)
    '
    '## Reset Event Label
    '
    pShowEvent ""
    '
    '## Select all text in the selected Textbox control
    '
    pHiLite txtDialog(Index)
    Select Case Index
        '
        '## First Textbox
        '
        Case [Node Text]
            meFocus = [Node Text]       '## Capture textbox focus
            Select Case meMode
                Case [Add Node]
                    '
                    '## Give user instructions
                    '
                    lblComments.Caption = "Enter the description of the new node."
                Case [Rename Node]
                    lblComments.Caption = "Select/Click the node to be renamed."
                Case [Move Node]
                    lblComments.Caption = "Select/Click the node to be moved."
                Case [Copy Node]                                                    '@@ v01.00.03
                    lblComments.Caption = "Select/Click the node to be copied."     '@@
                Case [Delete Node]
                    lblComments.Caption = "Select/Click the node to be deleted."
            End Select
        '
        '## Second Textbox
        '
        Case [Parent Node]
            meFocus = [Parent Node]     '## Capture textbox focus
            Select Case meMode
                Case [Add Node]
                    lblComments.Caption = "Select/Click the parent node (If not a root node)."
                Case [Rename Node]
                    lblComments.Caption = "Enter the new description of the selected node."
                Case [Move Node], [Copy Node]                                       '@@ v01.00.03
                    lblComments.Caption = "Select/Click the destination node."
            End Select

    End Select
End Sub

Private Sub txtDialog_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        '
        '## First Textbox
        '
        Case [Node Text]
            Select Case meMode
                Case [Rename Node], [Move Node], [Copy Node], [Delete Node] '@@ v01.00.03
'                Case [Rename Node], [Move Node], [Delete Node]
                    '
                    '## We only want the user to click on a node
                    '
                    KeyAscii = 0
                    Beep
            End Select
        '
        '## Second Textbox
        '
        Case [Parent Node]
            Select Case meMode
                Case [Add Node], [Move Node], [Copy Node]                   '@@ v01.00.03
'                Case [Add Node], [Move Node]
                    '
                    '## We only want the user to click on a node
                    '
                    KeyAscii = 0
                    Beep
            End Select

    End Select
End Sub

Private Sub txtDialog_LostFocus(Index As Integer)
    lblComments.Caption = ""
End Sub

'===========================================================================
' Form Events: Find controls
'
Private Sub cmdFind_Click(Index As Integer)

    Dim oBusy As cHourglass

    Set oBusy = New cHourglass
    Select Case Index
        Case [Find First]
            If Not pFindFirst Then
                Set oBusy = Nothing
                MsgBox "No record found.", _
                       vbApplicationModal + vbInformation + vbOKOnly, _
                       "Find Record"
                txtFind.SetFocus
            Else
                tvwDialog.SetFocus
            End If

        Case [Find Next]
            If Not pFindNext Then
                If mbFindNext Then
                    Set oBusy = Nothing
                    MsgBox "No further records found.", _
                           vbApplicationModal + vbInformation + vbOKOnly, _
                           "Find Record"
                Else
                    Set oBusy = Nothing
                    MsgBox "No record found.", _
                           vbApplicationModal + vbInformation + vbOKOnly, _
                           "Find Record"
                End If
                txtFind.SetFocus
            Else
                tvwDialog.SetFocus
            End If

        Case [Find Previous]
            If Not pFindPrevious Then
                If mbFindNext Then
                    Set oBusy = Nothing
                    MsgBox "No further records found.", _
                           vbApplicationModal + vbInformation + vbOKOnly, _
                           "Find Record"
                Else
                    Set oBusy = Nothing
                    MsgBox "No record found.", _
                           vbApplicationModal + vbInformation + vbOKOnly, _
                           "Find Record"
                End If
                txtFind.SetFocus
            Else
                tvwDialog.SetFocus
            End If

    End Select

End Sub

Private Sub optFind_Click(Index As Integer)
    meFindMode = CByte(Abs(optFind([Product]).Value = True))
    pFindReset
End Sub

Private Sub txtFind_Change()
    pFindReset
    cmdFind([Find First]).Enabled = (Len(txtFind.Text) > 0)
End Sub

Private Sub txtFind_GotFocus()
    pHiLite txtFind
End Sub

'===========================================================================
' cTreeView Class Events
'
Private Sub moTree_CopyNode(DestNode As MSComctlLib.Node, SrcNode As MSComctlLib.Node)  '@@ v01.00.03
    '
    '## Raised by moTree.NodeCopy to do the physical operation. Cannot be done from
    '   within the class due to too many external factors. In this case the key
    '   is the same one used in the database with a prefix - therefore needs to
    '   be generated external to the cTreeView class.
    '
    Dim eAddType As eFindMode
    Dim lID      As Long
    Dim oNode    As MSComctlLib.Node

    Select Case Left$(SrcNode.Key, 1)                   '## Extract node type
        Case "N": eAddType = [Group]
        Case "P": eAddType = [Product]
    End Select
    txtDialog([Node Text]).Text = SrcNode.Text          '## Set AddNode Text
    lID = pAddRecord(eAddType, DestNode)                '## Add the SrcNode
    If lID Then                                         '## Did we add it successfully to the Database?
        Set oNode = pAddNode(eAddType, SrcNode.Text, lID, DestNode) '## Yes. Add Node to TreeView
        If Len(msCopyKey) = 0 Then                      '## Have we caputed the first copied node's key?
            msCopyKey = oNode.Key                       '## No. Better do it then. (Used at end to select)
        End If
    Else
        moTree.CancelCopy = True                        '## No. Cancel copying nodes due to DB error!!
    End If

End Sub

Private Sub moTree_ContextMenu(Node As MSComctlLib.Node, x As Single, y As Single)
    '
    '## Right Button was pressed requesting a Context/Popup menu
    '
    Debug.Print ">> "; Me.Name; ".moTree::ContextMenu -> Node.Text = ";

    pShowNodeDetails Node
    pShowEvent "ContextMenu"
    If Not (Node Is Nothing) Then
        '
        '## Show popup menu for the specific node
        '
        Me.PopupMenu mnuPopNode, _
                     vbPopupMenuLeftAlign + vbPopupMenuRightButton, _
                     tvwDialog.Left + x, tvwDialog.Top + y                  '@@ v01.00.01
        Debug.Print "'"; Node.Text; "'"
    Else
        '
        '## Background clicked instead of node. Show popup menu for TreeView and not a node
        '
        Me.PopupMenu mnuPopTree, _
                     vbPopupMenuLeftAlign + vbPopupMenuRightButton, _
                     tvwDialog.Left + x, tvwDialog.Top + y                  '@@ v01.00.01
        Debug.Print "[Control Menu]"
        With mnuPopTree                             '@@ v01.00.03
            .Visible = True                         '@@ Required if popup menu is a child of
            .Enabled = True                         '@@ another popup menu
        End With                                    '@@
    End If
End Sub

Private Sub moTree_StartDrag(SourceNode As MSComctlLib.Node)
    '
    '## We've started dragging a node
    '
    Debug.Print "++ Start Drag Node = '"; SourceNode.Text; "'"

    pShowNodeDetails SourceNode
    pShowEvent "StartDrag"

End Sub

Private Sub moTree_Dragging(SourceNode As MSComctlLib.Node, TargetParent As MSComctlLib.Node)
    '
    '## Node being dragged
    '
    If msDragTarget <> TargetParent.Text Then   '## Only proceed if a different node
        Debug.Print "++ Dragging Node = '"; SourceNode.Text; "' with target node = '"; TargetParent.Text; "'"
        msDragTarget = TargetParent.Text
        If Left$(SourceNode.Key, 1) = "N" Then                      '@@ v01.00.01
            tvwDialog.DragIcon = imgDialog.ListImages(5).Picture    '@@ Set drag icon to type of node
        Else                                                        '@@
            tvwDialog.DragIcon = imgDialog.ListImages(8).Picture    '@@ Note: This overrides cTREEVIEW
        End If                                                      '@@       Class DragIcon
        pShowEvent "Dragging"
    End If

End Sub

Private Sub moTree_Dropped(SourceNode As MSComctlLib.Node, TargetParent As MSComctlLib.Node)
    '
    '## Node has been dropped - Now what to do with it...
    '
    Debug.Print "++ Dropped Node = '"; SourceNode.Text; "'"
    '
    '## Move the dragged node
    '
    pShowEvent "Dropped"
    'If Not moTree.NodeMove(TargetParent, SourceNode) Then
    If Not pMoveRecord(SourceNode, TargetParent) Then           '!! ADO Code
        '
        '## Problems with moving the node. Most likely a root node was dragged!
        '
        MsgBox "Unable to move the selected node.", _
               vbApplicationModal + vbExclamation + vbOKOnly, _
               App.Title
    End If
    pShowNodeDetails SourceNode

End Sub

Private Sub moTree_Selected(Node As MSComctlLib.Node)
    '
    '## A Node has been selected
    '
    With Node
        If InStr(.Key, "P") Then
            Debug.Print "## Product = [" + .Text + "]", "PkID = [" + Mid$(.Key, 2) + "]"
        Else
            Debug.Print "## Node Click = [" + .Text + "]"
        End If
    End With

    pShowNodeDetails Node
    pShowEvent "Selected"
    '
    '## Pass the text of the selected node to the correct Textbox based on the selected action
    '
    Select Case meFocus
        Case [Node Text]
            Select Case meMode
                Case [Rename Node], [Move Node], [Copy Node], [Delete Node] '@@ v01.00.03
'                Case [Rename Node], [Move Node], [Delete Node]
                    txtDialog([Node Text]).Text = Node.Text
                    With txtDialog([Parent Node])
                        '
                        '## Set the focus to the next control
                        '
                        If .Visible Then
                            .SetFocus
                        Else
                            cmdDialog([Execute Mode]).SetFocus
                        End If
                    End With
                    '
                    '## Store the selected node
                    '
                    Set moSelectedNode = Node
                    If Not (meMode = [Delete Node]) Then
                        meFocus = [Parent Node]
                    End If
            End Select

        Case [Parent Node]
            Select Case meMode
                Case [Add Node], [Move Node], [Copy Node]                   '@@ v01.00.03
'                Case [Add Node], [Move Node]
                    txtDialog([Parent Node]).Text = Node.Text
                    cmdDialog([Execute Mode]).SetFocus
'                    If meMode = [Move Node] Then
                    Select Case meMode                                      '@@ v01.00.03
                        Case [Move Node], [Copy Node]                       '@@
                            Set moDestNode = Node
                        Case Else
                            Set moSelectedNode = Node
                    End Select
'                    Else
'                        Set moSelectedNode = Node
'                    End If
            End Select
    End Select

End Sub

'===========================================================================
' Private subroutines and functions
'
Private Sub Action(State As Integer)

    Dim sText As String
    Dim oNode As MSComctlLib.Node
    Dim eType As eFindMode                                      '@@ v01.00.03
    Dim lID   As Long                                           '@@

    Select Case State
        '
        '## Setup user frame and contained controls based on action
        '
        Case [Add Node]
            meMode = [Add Node]
            fraDialog.Caption = "Add Node:"
            With lblDialog([Node Text])
                .Caption = "Node Text: "
                .Visible = True
            End With
            With txtDialog([Node Text])
                .Text = ""
                .Visible = True
                .SetFocus
            End With
            With lblDialog([Parent Node])
                .Caption = "Parent Node: "
                .Visible = True
            End With
            With txtDialog([Parent Node])
                .Text = ""
                .Visible = True
            End With
            With cmdDialog([Execute Mode])
                .Top = lblDialog([Parent Node]).Top
                .Visible = True
            End With
            With chkDialog([Action Option])             '@@ v01.00.03
                .Caption = "Is Product?"                '@@
                .Visible = True                         '@@
            End With

        Case [Rename Node]
            meMode = [Rename Node]
            fraDialog.Caption = "Rename Node:"
            With lblDialog([Node Text])
                .Caption = "Old Node Text: "
                .Visible = True
            End With
            With txtDialog([Node Text])
                .Text = ""
                .Visible = True
                .SetFocus
            End With
            With lblDialog([Parent Node])
                .Caption = "New Node Text: "
                .Visible = True
            End With
            With txtDialog([Parent Node])
                .Text = ""
                .Visible = True
            End With
            With cmdDialog([Execute Mode])
                .Top = lblDialog([Parent Node]).Top
                .Visible = True
            End With
            With chkDialog([Action Option])             '@@ v01.00.03
                .Caption = ""                           '@@
                .Visible = False                        '@@
            End With

        Case [Move Node]
            meMode = [Move Node]
            fraDialog.Caption = "Move Node:"
            With lblDialog([Node Text])
                .Caption = "From Node: "
                .Visible = True
            End With
            With txtDialog([Node Text])
                .Text = ""
                .Visible = True
                .SetFocus
            End With
            With lblDialog([Parent Node])
                .Caption = "To Node: "
                .Visible = True
            End With
            With txtDialog([Parent Node])
                .Text = ""
                .Visible = True
            End With
            With cmdDialog([Execute Mode])
                .Top = lblDialog([Parent Node]).Top
                .Visible = True
            End With
            With chkDialog([Action Option])             '@@ v01.00.03
                .Caption = ""                           '@@
                .Visible = False                        '@@
            End With

        '@@ v01.00.03 ------ Start ----------------
        Case [Copy Node]
            meMode = [Copy Node]
            fraDialog.Caption = "Copy Node:"
            With lblDialog([Node Text])
                .Caption = "From Node: "
                .Visible = True
            End With
            With txtDialog([Node Text])
                .Text = ""
                .Visible = True
                .SetFocus
            End With
            With lblDialog([Parent Node])
                .Caption = "To Node: "
                .Visible = True
            End With
            With txtDialog([Parent Node])
                .Text = ""
                .Visible = True
            End With
            With cmdDialog([Execute Mode])
                .Top = lblDialog([Parent Node]).Top
                .Visible = True
            End With
            With chkDialog([Action Option])
                .Caption = "Children?"
                .Visible = True
            End With
        '@@ v01.00.03 ------ Finish ---------------

        Case [Delete Node]
            meMode = [Delete Node]
            fraDialog.Caption = "Delete Node:"
            With lblDialog([Node Text])
                .Caption = "Node: "
                .Visible = True
            End With
            With txtDialog([Node Text])
                .Text = ""
                .Visible = True
                .SetFocus
            End With
            lblDialog([Parent Node]).Visible = False
            txtDialog([Parent Node]).Visible = False
            With cmdDialog([Execute Mode])
                .Top = lblDialog([Node Text]).Top
                .Visible = True
            End With
            With chkDialog([Action Option])             '@@ v01.00.03
                .Caption = ""                           '@@
                .Visible = False                        '@@
            End With

        '
        '## Time to put the wheels in motion
        '
        Case [Execute Mode]
            Select Case meMode
                Case [Add Node]
                    On Error Resume Next
                    If Len(txtDialog([Node Text]).Text) Then
                        sText = txtDialog([Node Text]).Text
                        eType = CByte(Abs(chkDialog([Action Option]).Value = 1))   '@@ v01.00.03
                        If Len(txtDialog([Parent Node]).Text) Then
                            '
                            '## If a destination node then node will be a child
                            '
'                            If (moTree.NodeFind(oNode, sText)) Then
'                                MsgBox "Node already exists.", _
'                                       vbApplicationModal + vbExclamation + vbOKOnly, _
'                                       "Add Node"
'                            Else
                                lID = pAddRecord(eType, moSelectedNode)         '@@ v01.00.03
                                If lID Then                                     '@@
                                    pAddNode eType, sText, lID, moSelectedNode
                                Else                                            '@@
                                    '!! ERROR
                                End If                                          '@@
'                            End If
                        Else
                            '
                            '## No Destination means node will be a root node
                            '
                            If (moTree.NodeFind(oNode, sText, sText)) Then
                                MsgBox "Node already exists.", _
                                       vbApplicationModal + vbExclamation + vbOKOnly, _
                                       "Add Node"
                            Else
                                If eType = [Group] Then                         '@@ v01.00.03
                                    lID = pAddRecord(eType)                     '@@
                                    If lID Then                                 '@@
                                        pAddNode eType, sText, lID              '@@
                                    Else                                        '@@
                                        '!! ERROR
                                    End If
                                Else                                            '@@
                                    MsgBox "Product requires a Parent Node.", _
                                           vbApplicationModal + vbExclamation + vbOKOnly, _
                                           "Add Node"                           '@@
                                End If                                          '@@
                            End If
                        End If
                        '
                        '## NodeAdd had a problem
                        '
                        If Err.Number Then
                            MsgBox "Error adding node.", _
                                   vbApplicationModal + vbExclamation + vbOKOnly, _
                                   "Add Node"
                            Err.Number = 0
                            Set moSelectedNode = Nothing
                            txtDialog([Parent Node]).Text = ""
                            txtDialog([Node Text]).SetFocus
                        Else
                            pShowEvent "[Added]"
                            tvwDialog.SetFocus
                        End If
                    Else
                        '
                        '## No text was entered for the new node
                        '
                        MsgBox "No new text entered.", _
                               vbApplicationModal + vbExclamation + vbOKOnly, _
                               "Add Node"
                        txtDialog([Node Text]).SetFocus
                    End If

                Case [Rename Node]
                    If Len(txtDialog([Node Text]).Text) Then
                        If Len(txtDialog([Parent Node]).Text) Then
                            'If Not moTree.NodeRename(moSelectedNode, txtDialog([Parent Node]).Text, True) Then
                            If Not pRenameRecord(moSelectedNode, txtDialog([Parent Node]).Text) Then    '!! ADO Code
                                '
                                '## Problem with renaming the node
                                '
                                MsgBox "Unable to rename the selected node.", _
                                       vbApplicationModal + vbExclamation + vbOKOnly, _
                                       "Rename Node"
                                txtDialog([Parent Node]).SetFocus
                            Else
                                pShowEvent "[Renamed]"
                                tvwDialog.SetFocus
                            End If
                        Else
                            '
                            '## No text was entered
                            '
                            MsgBox "No new text entered.", _
                                   vbApplicationModal + vbExclamation + vbOKOnly, _
                                   "Rename Node"
                            txtDialog([Parent Node]).SetFocus
                        End If
                    Else
                        '
                        '## No node was selected
                        '
                        MsgBox "No node to be renamed was selected.", _
                               vbApplicationModal + vbExclamation + vbOKOnly, _
                               "Rename Node"
                        txtDialog([Node Text]).SetFocus
                    End If

                Case [Move Node]
                    If Len(txtDialog([Node Text]).Text) Then
                        If Len(txtDialog([Parent Node]).Text) Then
                            'If Not moTree.NodeMove(moDestNode, moSelectedNode, True) Then
                            If Not pMoveRecord(moSelectedNode, moDestNode) Then     '!! ADO Code
                                '
                                '## Problem moving the node. Most likely a root node was selected.
                                '
                                MsgBox "Unable to move the selected node.", _
                                       vbApplicationModal + vbExclamation + vbOKOnly, _
                                       "Move Node"
                                txtDialog([Node Text]).SetFocus
                            Else
                                pShowEvent "[Moved]"
                                tvwDialog.SetFocus
                            End If
                        Else
                            '
                            '## No destination node selected
                            '
                            MsgBox "No Destination node selected.", _
                                   vbApplicationModal + vbExclamation + vbOKOnly, _
                                   "Move Node"
                            txtDialog([Parent Node]).SetFocus
                        End If
                    Else
                        '
                        '## No source node selected
                        '
                        MsgBox "No Source node selected.", _
                               vbApplicationModal + vbExclamation + vbOKOnly, _
                               "Move Node"
                        txtDialog([Node Text]).SetFocus
                    End If

                '@@ v01.00.03 ------ Start ----------------
                Case [Copy Node]
                    Dim bAction As Boolean
                    Dim oBusy As cHourglass
                    Set oBusy = New cHourglass                  '## Long operation if lots of nodes to
                                                                '   copy. Therefore let the user know
                                                                '   something is happening
                    If Len(txtDialog([Node Text]).Text) Then
                        If Len(txtDialog([Parent Node]).Text) Then
                            tvwDialog.Visible = False           '## Speed up drawing and stop flickering
                            bAction = CBool(chkDialog([Action Option]).Value = 1)
                            Dim oTmr As cBenchmark
                            Set oTmr = New cBenchmark
                            oTmr.Start                          '## Start timing this action
                            '
                            '## Did we copy all required nodes?
                            '
                            If Not moTree.NodeCopy(moDestNode, moSelectedNode, bAction) Then
                                '
                                '## No. Problem copying the node. Most likely an ADO error.
                                '
                                MsgBox "Unable to copy the selected node.", _
                                       vbApplicationModal + vbExclamation + vbOKOnly, _
                                       "Copy Node"
                                txtDialog([Node Text]).SetFocus
                                tvwDialog.Visible = True
                            Else
                                '
                                '## Yes. Cleanup by collapsing the source node, restore
                                '   the from node text, & lastly expand and select the
                                '  copied node
                                '
                                moSelectedNode.Expanded = False
                                txtDialog([Node Text]).Text = moSelectedNode.Text
                                With tvwDialog.Nodes(msCopyKey)
                                    If Left$(.Key, 1) = "N" Then
                                        .Expanded = False
                                    End If
                                    If bAction Then
                                        .Expanded = True
                                    End If
                                    .EnsureVisible
                                    .Selected = True
                                End With
                                pShowEvent "[Copied]"
                                tvwDialog.Visible = True
                                tvwDialog.SetFocus
                            End If
                            oTmr.Finish
                            pShowTimer oTmr.ElapsedTime
                        Else
                            '
                            '## No destination node selected
                            '
                            MsgBox "No Destination node selected.", _
                                   vbApplicationModal + vbExclamation + vbOKOnly, _
                                   "Copy Node"
                            txtDialog([Parent Node]).SetFocus
                        End If
                    Else
                        '
                        '## No source node selected
                        '
                        MsgBox "No Source node selected.", _
                               vbApplicationModal + vbExclamation + vbOKOnly, _
                               "Copy Node"
                        txtDialog([Node Text]).SetFocus
                    End If
                    msCopyKey = ""              '## Reset ket tag until next operation
                '@@ v01.00.03 ------ Finish ---------------

                Case [Delete Node]
                    If Len(txtDialog([Node Text]).Text) Then
                        If (MsgBox("Are you sure?", _
                                    vbApplicationModal + vbDefaultButton2 + vbQuestion + vbYesNo, _
                                    "Delete Node: " + moSelectedNode.Text) = vbYes) Then                   '!! ADO Code version
'                        If (MsgBox("Child Nodes will also be deleted. Are you sure?", _
'                                    vbApplicationModal + vbDefaultButton2 + vbQuestion + vbYesNo, _
'                                    "Delete Node") = vbYes) Then
                            'If Not moTree.NodeDelete(moSelectedNode, True) Then
                            If Not pDeleteRecord(moSelectedNode) Then               '!! ADO Code
                                '
                                '## Problem deleting the selected node. Most Likely a root node was selected.
                                '
                                MsgBox "Unable to delete the selected node.", _
                                       vbApplicationModal + vbExclamation + vbOKOnly, _
                                       "Delete Node"
                                txtDialog([Node Text]).SetFocus
                            Else
                                pShowEvent "[Deleted]"
                                tvwDialog.SetFocus
                            End If
                        End If
                    Else
                        '
                        '## No source node selected
                        '
                        MsgBox "No Source node selected.", _
                               vbApplicationModal + vbExclamation + vbOKOnly, _
                               "Move Node"
                        txtDialog([Node Text]).SetFocus
                    End If

            End Select
    End Select

End Sub

Private Sub pHiLite(txtBox As TextBox)
    '
    '## Selects all text of the designated Textbox
    '
    With txtBox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub pShowEvent(sText As String)
    '
    '## Show event to user
    '
    lblEvent(1).Caption = sText
End Sub

Private Sub pShowTimer(dElapsed As Double)
    '
    '## Show event to user
    '
    lblEvent(3).Caption = Format$(dElapsed, "#,##0.00000     ")
End Sub

Private Sub pShowNodeDetails(Node As MSComctlLib.Node)
    '
    '## Show all node details to user
    '
    Dim sText As String

    If Node Is Nothing Then Exit Sub
    With Node
        lblDetails(0).Caption = .Text
        lblDetails(1).Caption = "&&H" + Hex$(.BackColor)
        lblDetails(2).Caption = IIf(.Bold, "Yes", "No")
        lblDetails(3).Caption = IIf(.Checked, "Yes", "No")
        If .Children Then
            sText = "Yes, " + CStr(.Children)
        Else
            sText = "No"
        End If
        lblDetails(4).Caption = sText
        lblDetails(5).Caption = IIf(.Expanded, "Yes", "No")
        lblDetails(6).Caption = "&&H" + Hex$(.ForeColor)
        lblDetails(7).Caption = "  \" + .FullPath
        lblDetails(8).Caption = CStr(.Index)
        lblDetails(9).Caption = .Key
        If moTree.IsRootNode(Node) Then
            sText = "Node is Root"
        Else
            sText = .Parent.Text
        End If
        lblDetails(10).Caption = sText
        lblDetails(11).Caption = .Root.Text
        lblDetails(12).Caption = IIf(.Selected, "Yes", "No")
        lblDetails(13).Caption = IIf(.Sorted, "Yes", "No")
        lblDetails(14).Caption = IIf(.Visible, "Yes", "No")
    End With

End Sub

'===========================================================================
' Private ADO subroutines and functions
'
Private Function pExpandNode(Node As MSComctlLib.Node) As Boolean

    If mbIsBusy Then Exit Function

    With Node
        If InStr(.Key, "N") And Len(.Tag) = 0 Then      '## Is a group node & is it loaded?
            If .Children Then                           '## No. Is there a dummy node?      '@@ v01.00.03
                tvwDialog.Nodes.Remove .Child.Index     '## Yes. Delete dummy node
                pLoadGroups Mid$(.Key, 2)               '## Load child groups
                pLoadProducts Node                      '## Load child products
                .Tag = "Loaded"                         '## Tag node as loaded
            End If
        End If
    End With

End Function

Private Sub pInitData()
    '
    '## Initialise database. NOTE: DATABASE MUST BE IN APP PATH.
    '
    mbIsDirty = True
    Set mcDB = New cDB
    mcDB.InitDB App.Path + "\Grouping.mdb"
    pLoadGroups
    mbIsDirty = False

End Sub

Private Sub pLoadGroups(Optional GroupID As Long = 0)

    Dim lLoop  As Long
    Dim lCount As Long
    Dim oRs    As ADODB.Recordset
    Dim oNode  As MSComctlLib.Node
    Dim sKey   As String                '## Key Prefixes: "N" = Group; "P" = Product; "D" = Dummy

    mbIsBusy = True
    '
    '## Did we successfully create the recordset?
    '
    If mcDB.CreateRS(oRs, mcSQL_GRP1 + CStr(GroupID) + mcSQL_GRP2) Then
        '
        '## Yes.
        '
        lCount = mcDB.RecordCount(oRs)
        If lCount Then
            moTree.Redraw False
            With oRs
                '
                '## For every branch
                '
                For lLoop = 1 To lCount
                    sKey = "N" + CStr(!PkID)
                    If !GroupID Then
                        '
                        '## It's a Child Node
                        '@@ v 01.00.01 Added icon pointers
                        '
                        moTree.NodeAdd "N" + CStr(!GroupID), tvwChild, sKey, !Desc, 3, 3, , True, , , False, , , , , 4
                        '
                        '## Add dummy key to Group node to display the Expand Handle (plus sign)
                        '
                        moTree.NodeAdd sKey, tvwChild, "D" + sKey, "Dummy", , , , , , , False
                    Else
                        '
                        '## It's a Root Node
                        '@@ v 01.00.01 Added icon pointers
                        '
                        moTree.NodeAdd , , sKey, !Desc, 1, 1, , True, , , False, , , , , 2
                        '
                        '## Add dummy key to Group node to display the Expand Handle (plus sign)
                        '
                        moTree.NodeAdd sKey, tvwChild, "D" + sKey, "Dummy", , , , , , , False
                    End If
                    '
                    '## Next node/record
                    '
                    mcDB.MoveDB emrMoveNext, oRs
                Next
            End With
            moTree.Redraw True
        End If
    End If
    mbIsBusy = False

End Sub

Private Sub pLoadProducts(oNode As Node)

    Dim lLoop  As Long
    Dim lCount As Long
    Dim lPtr   As Long
    Dim oRs    As ADODB.Recordset

    With oNode
        Debug.Print "** Loading Products for Node [" + oNode.Text + "]"
        If mcDB.CreateRS(oRs, mcSQL_PROD1 + Mid$(.Key, 2) + mcSQL_PROD2) Then
            lCount = mcDB.RecordCount(oRs)
            If lCount Then
                For lLoop = 1 To lCount
                    '
                    '## Add product nodes to group node
                    '@@ v 01.00.01 Added icon pointers
                    '
                    moTree.NodeAdd .Key, tvwChild, "P" + CStr(oRs!PkID), oRs!Desc, 6, 6, , , , , , , , clPRODCOLOR, , 6
                    '
                    '## Next node/record
                    '
                    mcDB.MoveDB emrMoveNext, oRs
                Next
            End If
        End If
    End With

End Sub

Private Function pAddNode(AddType As eFindMode, _
                          NewText As String, _
                          ID As Long, _
                 Optional DestNode As MSComctlLib.Node) As MSComctlLib.Node '@@ v01.00.03

    Dim sKey  As String
    Dim lIcon As Long
    Dim lColr As Long
    Dim bBold As Boolean

    Select Case AddType
        Case [Group]
            sKey = "N"
            lIcon = 3
            bBold = True
            lColr = vbButtonText
        Case [Product]
            sKey = "P"
            lIcon = 6
            bBold = False
            lColr = clPRODCOLOR
    End Select

    If DestNode Is Nothing Then
        If AddType = [Group] Then
            Set pAddNode = moTree.NodeAdd(, , sKey + CStr(ID), NewText, 1, 1, , _
                                          bBold, , , , True, , lColr, , 2)
        End If
    Else
        DestNode.Expanded = True
        Set pAddNode = moTree.NodeAdd(DestNode.Key, tvwChild, sKey + CStr(ID), NewText, lIcon, lIcon, , _
                                      bBold, , True, , , , lColr, , lIcon + 1)
    End If

End Function

Private Function pAddRecord(AddType As eFindMode, _
                   Optional ToNode As MSComctlLib.Node) As Long     '@@ v01.00.03
    '
    '## NOTE: I've used INSERT INTO with Parameters through a Command.Exectue
    '         instead of Recordset.AddNew. This is much faster and uses far
    '         less resources when working with large amounts of data.
    '
    Dim oRs      As ADODB.Recordset
    Dim oCmd     As ADODB.Command
    Dim lGroupID As Long

    Set oCmd = New ADODB.Command

    On Error GoTo ErrorHandler

    pExpandNode ToNode
    Select Case AddType
        Case [Group]
            With oCmd
                .CommandType = adCmdText
                .CommandText = mcSQL_AGRP
                .Parameters.Append .CreateParameter("Desc", adVarWChar, adParamInput, 30, txtDialog([Node Text]).Text)
                If ToNode Is Nothing Then
                    lGroupID = 0
                Else
                    lGroupID = CLng(Mid$(ToNode.Key, 2))
                End If
                .Parameters.Append .CreateParameter("GroupID", adInteger, adParamInput, , lGroupID)
                .Parameters.Append .CreateParameter("Type", adInteger, adParamInput, , 0)
                .Parameters.Append .CreateParameter("Active", adBoolean, adParamInput, , True)
            End With
            If mcDB.ExecuteSQL(, oCmd) Then
                pAddRecord = mcDB.NewID
            End If

        Case [Product]
            If Not (ToNode Is Nothing) Then
                With oCmd
                    .CommandType = adCmdText
                    .CommandText = mcSQL_APRD
                    .Parameters.Append .CreateParameter("Desc", adVarWChar, adParamInput, 30, txtDialog([Node Text]).Text)
                    .Parameters.Append .CreateParameter("GroupID", adInteger, adParamInput, , CLng(Mid$(ToNode.Key, 2)))
                    .Parameters.Append .CreateParameter("Active", adBoolean, adParamInput, , True)
                End With
                If mcDB.ExecuteSQL(, oCmd) Then
                    pAddRecord = mcDB.NewID
                End If
            End If
    End Select

ErrorHandler:
    '## Failed Operation
End Function


Private Function pDeleteRecord(Node As MSComctlLib.Node) As Boolean

    Dim SSQL   As String
    Dim lCount As Long
    Dim lLoop  As Long
    Dim oNode As MSComctlLib.Node

    With Node
        '
        '## Make sure that we're not working with the root node
        '
        If Not moTree.IsRootNode(Node) Then
            moTree.Redraw True
            lCount = .Children
            Set oNode = .Child
            If lCount Then
                '
                '## If there are child nodes, then we must not have broken
                '   links. Set children's parent to the parent of the
                '   node being deleted
                '
                For lLoop = 1 To lCount
                    If Not pMoveRecord(oNode, .Parent) Then
                        '
                        '## Big problem - This shouldn't happen!
                        '
                        MsgBox "Error whilst moving node '" + oNode.Text + "' to node '" + .Parent.Text + "'.", _
                               vbApplicationModal + vbCritical + vbOKOnly, _
                               "Delete Record '" + .Text + "'."
                        moTree.Redraw True
                        Exit Function
                    End If
                    If lLoop < lCount Then
                        '
                        '## Select next child. Note: Next child becomes the first child after the
                        '   first child is deleted.
                        '
                        Set oNode = .Child
                    End If
                Next
            End If
            '
            '## Now that the node's isolated, delete the node from the tree and the DB
            '
            If moTree.NodeDelete(Node, True) Then
                Select Case Left$(Node.Key, 1)
                    Case "N"
                        '
                        '## Build Group Delete command
                        '
                        SSQL = mcSQL_DGRP1 + Mid$(Node.Key, 2) + mcSQL_DGRP2
                    Case "P"
                        '
                        '## Build Product Delete command
                        '
                        SSQL = mcSQL_DPRD1 + Mid$(Node.Key, 2) + mcSQL_DPRD2
                End Select
                '
                '## Delete Record
                '
                pDeleteRecord = mcDB.ExecuteSQL(SSQL)
                pFindReset                          '## Recordset needs to be rebuilt
            End If
            moTree.Redraw False
        End If
    End With

End Function

Private Function pMoveRecord(Node As MSComctlLib.Node, _
                             ToNode As MSComctlLib.Node) As Boolean
    Dim SSQL As String

    If Left$(ToNode.Key, 1) = "P" Then
        '
        '## Cannot move a group or product onto a product. Therefore
        '    move to product's group (Parent Node).
        '
        Set ToNode = ToNode.Parent
    End If

    If Not moTree.IsParentNode(ToNode, Node) Then
        '
        '## We don't want to be able to move a parent to & below its child node.
        '
        If moTree.NodeMove(ToNode, Node, True) Then
            Select Case Left$(Node.Key, 1)
                Case "N"
                    '
                    '## Build Group Link ID
                    '
                    SSQL = mcSQL_UGRP1 + Mid$(ToNode.Key, 2) + mcSQL_UGRP2 + Mid$(Node.Key, 2) + mcSQL_UGRP3
                Case "P"
                    '
                    '## Build Product Link ID
                    '
                    SSQL = mcSQL_UPRD1 + Mid$(ToNode.Key, 2) + mcSQL_UPRD2 + Mid$(Node.Key, 2) + mcSQL_UPRD3
            End Select
            '
            '## Update Record
            '
            pMoveRecord = mcDB.ExecuteSQL(SSQL)
            pFindReset                              '## Recordset needs to be rebuilt
        End If
    End If

End Function

Private Function pRenameRecord(Node As MSComctlLib.Node, NewText As String) As Boolean

    Dim SSQL   As String

    If moTree.NodeRename(Node, NewText, True) Then
        Select Case Left$(Node.Key, 1)
            Case "N"
                '
                '## Build Group Rename command
                '
                SSQL = mcSQL_RGRP1 + NewText + mcSQL_RGRP2 + Mid$(Node.Key, 2) + mcSQL_RGRP3
            Case "P"
                '
                '## Build Product Rename command
                '
                SSQL = mcSQL_UPRD1 + NewText + mcSQL_UPRD2 + Mid$(Node.Key, 2) + mcSQL_UPRD3
        End Select
        '
        '## Update Record
        '
        pRenameRecord = mcDB.ExecuteSQL(SSQL)
        pFindReset                                  '## Recordset needs to be rebuilt
    End If

End Function

'===========================================================================
' Private ADO subroutines and functions : Find First/Next Product/Group
'
Private Function pFindFirst() As Boolean

    Dim SSQL As String

    '
    '## Build Query
    '
    Select Case meFindMode
        Case [Group]:   SSQL = mcSQL_FGRP
        Case [Product]: SSQL = mcSQL_FPRD
    End Select

    Dim oTmr As cBenchmark                                  '@@ v01.00.03
    Set oTmr = New cBenchmark                               '@@
    oTmr.Start                                              '@@

    With mcDB
        '
        '## Retrieve records
        '
        If .CreateRS(moFindRS, SSQL) Then
            '
            '## Do we have any records?
            '
            If .RecordCount(moFindRS) Then
                '
                '## Yes. Now filter
                '
                moFindRS.Filter = "Desc Like '*" + txtFind.Text + "*'"
                '
                '## Do we have any records left?
                '
                If .RecordCount(moFindRS) Then
                    '
                    '## Yes. Retrieve all groups.
                    '
                    If .CreateRS(moGrpRS, mcSQL_FGRP) Then
                        '
                        '## Do we have any records?
                        '
                        If .RecordCount(moGrpRS) Then
                            '
                            '## Yes. Now find the first record
                            '
                            .MoveDB emrMoveFirst, moFindRS
                            pFindFirst = pFindNext
                        End If
                    End If
                End If
            End If
        End If
    End With
    oTmr.Finish                                             '@@ v01.00.03
    pShowTimer oTmr.ElapsedTime                             '@@
    pShowEvent "*Find First*"

End Function

Private Function pFindNext() As Boolean

    Dim oNode     As MSComctlLib.Node
    Dim sText     As String
    Dim sKey      As String
    Dim vBookmark As Variant

    On Error Resume Next
    If mbFindNext Then
        '
        '## Now looking for the next record
        '
        Dim oTmr As cBenchmark                              '@@ v01.00.03
        Set oTmr = New cBenchmark                           '@@
        oTmr.Start                                          '@@
    
        vBookmark = moFindRS.Bookmark
        moFindRS.MoveNext
        '## NOTE: BOF wont work if recordset has a filter applied.
        '         Therefore capture the error.
        If Err.Number Then
            '## We shouldn't ever get here...
            moFindRS.Bookmark = vBookmark
            Exit Function
        End If
    End If
    sText = moFindRS!Desc
    On Error GoTo 0
    '
    '## Are we at the end of the filtered recordset?
    '
    If Len(sText) Then
        '
        '## No. Build Query
        '
        Select Case meFindMode
            Case [Group]:   sKey = "N" + CStr(moFindRS!PkID)
            Case [Product]: sKey = "P" + CStr(moFindRS!PkID)
        End Select
        '
        '## Is the node already loaded?
        '
        If moTree.NodeFind(oNode, sText, sKey, True) Then
            '
            'Yes
            '
            oNode.EnsureVisible
            'mbFindNext = True                               '@@ v01.00.01 Forgot to enable the
            'cmdFind([Find Next]).Enabled = True             '@@           next/previous buttons
            'cmdFind([Find Previous]).Enabled = True         '@@           here - oops!
            pFindNext = True
        Else
            '
            '## Back build parent branches. This routine will only fail if
            '   database table integritry error.
            '
            If pFindBackBuild(moFindRS!GroupID) Then
                '
                '## Find the node
                '
                If moTree.NodeFind(oNode, sText, sKey, True) Then
                    'mbFindNext = True
                    'cmdFind([Find Next]).Enabled = True
                    'cmdFind([Find Previous]).Enabled = True
                    pFindNext = True
                    '
                    '## Make sure we can see our search result
                    '
                    oNode.EnsureVisible
                Else
                    '
                    '## WARNING!! We have some kind of serious problem!!
                    '             Most likely a MS TreeView bug has put us here...
                    '
                End If
            End If
        End If
    Else
        '## Not found. Therefore point to last success.
        moFindRS.Bookmark = vBookmark
    End If
    If mbFindNext Then                                      '@@ v01.00.03
        oTmr.Finish                                         '@@
        pShowTimer oTmr.ElapsedTime                         '@@
        pShowEvent "*Find Next*"                            '@@
    ElseIf pFindNext Then                                   '@@
        mbFindNext = True                                   '@@
        cmdFind([Find Next]).Enabled = True                 '@@
        cmdFind([Find Previous]).Enabled = True             '@@
    End If                                                  '@@

End Function

Private Function pFindPrevious() As Boolean

    Dim oNode     As MSComctlLib.Node
    Dim sText     As String
    Dim sKey      As String
    Dim vBookmark As Variant

    Dim oTmr As cBenchmark                                  '@@ v01.00.03
    Set oTmr = New cBenchmark                               '@@
    oTmr.Start                                              '@@

    On Error Resume Next
    If mbFindNext Then
        vBookmark = moFindRS.Bookmark
        moFindRS.MovePrevious
        '## NOTE: BOF wont work if recordset has a filter applied.
        '         Therefore capture the error.
        If Err.Number Then
            '## We shouldn't ever get here...
            moFindRS.Bookmark = vBookmark
            Exit Function

        End If
    End If
    sText = moFindRS!Desc
    On Error GoTo 0
    If Len(sText) Then
        Select Case meFindMode
            Case [Group]:   sKey = "N" + CStr(moFindRS!PkID)
            Case [Product]: sKey = "P" + CStr(moFindRS!PkID)
        End Select
        '
        '## Find the node. We already know it's loaded.
        '
        If moTree.NodeFind(oNode, sText, sKey, True) Then
            pFindPrevious = True
        End If
    Else
        '## Not found. Therefore point to last success.
        moFindRS.Bookmark = vBookmark
    End If
    oTmr.Finish                                             '@@ v01.00.03
    pShowTimer oTmr.ElapsedTime                             '@@
    pShowEvent "*Find Previous*"                            '@@

End Function

Private Function pFindBackBuild(PkID As Long) As Boolean
    '
    '## This routine will recursively step back and load required
    '   branches if not already expanded.
    '
    Dim oNode   As MSComctlLib.Node
    Dim sText   As String
    Dim sKey    As String

    '
    '## Get Group Text
    '
    If mcDB.FindDB(efrFindFirst, "PkID=" + CStr(PkID), moGrpRS) Then
        sText = moGrpRS!Desc
        sKey = "N" + CStr(moGrpRS!PkID)
        '
        '## Is Group already loaded?
        '
        If Not moTree.NodeFind(oNode, sText, sKey, False) Then
            '
            '## No. Try its parent
            '
            pFindBackBuild = pFindBackBuild(moGrpRS!GroupID)
            '
            '## Point to the Group now that it's loaded
            '
            moTree.NodeFind oNode, sText, sKey, False
        End If
        If Not oNode Is Nothing Then    '@@ v01.00.03 - Extra error protection
            pFindBackBuild = True       '## We've had success.
            pExpandNode oNode           '## Expand and load Group
        End If
    Else
        '## WARNING!! Database integrity error!!
    End If

End Function

Private Sub pFindReset()
    '
    '## Disable Find Next & Previous buttons
    '
    mbFindNext = False
    cmdFind([Find Next]).Enabled = False
    cmdFind([Find Previous]).Enabled = False
End Sub
