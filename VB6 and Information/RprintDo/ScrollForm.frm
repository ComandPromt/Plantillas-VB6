VERSION 5.00
Begin VB.Form ListFrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ScrollFrm"
   ClientHeight    =   5685
   ClientLeft      =   1980
   ClientTop       =   2535
   ClientWidth     =   9450
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   9450
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Corner 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      DrawStyle       =   3  'Dash-Dot
      DrawWidth       =   2
      FillColor       =   &H000000FF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   435
      Left            =   10695
      ScaleHeight     =   435
      ScaleWidth      =   375
      TabIndex        =   9
      Top             =   7485
      Width           =   375
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      LargeChange     =   1000
      Left            =   15
      SmallChange     =   100
      TabIndex        =   2
      Top             =   5490
      Width           =   6300
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5295
      LargeChange     =   1000
      Left            =   7740
      SmallChange     =   100
      TabIndex        =   1
      Top             =   390
      Value           =   100
      Width           =   240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   9060
      Left            =   -480
      TabIndex        =   0
      Top             =   30
      Width           =   10290
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   225
         TabIndex        =   7
         Top             =   6720
         Width           =   2220
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         Columns         =   2
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3300
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   2820
         Width           =   8385
      End
      Begin VB.CheckBox Yes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Yes"
         Height          =   255
         Left            =   3195
         TabIndex        =   5
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox Not 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Not"
         Height          =   255
         Left            =   4920
         TabIndex        =   4
         Top             =   1785
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Some Performances of Roboprint"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   3
         Left            =   2745
         TabIndex        =   10
         Top             =   1095
         Width           =   5070
      End
      Begin VB.Shape Shape1 
         Height          =   555
         Left            =   2895
         Top             =   1725
         Width           =   3045
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "FileListBox"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   300
         TabIndex        =   8
         Top             =   6330
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Scroll Frame by Robocx"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   1
         Left            =   2775
         TabIndex        =   3
         Top             =   -15
         Width           =   3930
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Some Performances of Roboprint"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Index           =   0
         Left            =   2760
         TabIndex        =   11
         Top             =   1095
         Width           =   5070
      End
   End
End
Attribute VB_Name = "ListFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const inch = 1440 'Twips per Inch
Private Sub Form_Load()
Frame1.Move 0, 0, 8.5 * inch, 11 * inch
Scroll_Resize
  Dim i   As Integer
  
   List1.AddItem "FontSize"
   List1.AddItem "ForeColor"
   List1.AddItem "ListBox Checked"
   List1.AddItem "TextBox MultiPage"
   List1.AddItem "ListBox MultiPage"
   List1.AddItem "DBGrid MultiPage"
   List1.AddItem "DBGrid Expanded"
   List1.AddItem "MsFlexGrid MulTiPage"
   List1.AddItem "MsFlexGrid Expanded"
   List1.AddItem "MsFlexGrid MergeCells"
   List1.AddItem "ListView"
   List1.AddItem "RichText in VB6 Version"
   List1.AddItem "OptionButton Value"
   List1.AddItem "CheckBox Value"
   List1.AddItem "Zoom"
   List1.AddItem "Orientation"
   List1.AddItem "MaskEd Edit"
   List1.AddItem "MSChart"
   List1.AddItem "Label.BackStyle Transparent"
   For i = 0 To 18
   List1.Selected(i) = True
   Next i
    List1.Selected(4) = False
   ' List1.Selected(11) = False
    Yes.Value = 1
End Sub

Private Sub Scroll_Resize()
VScroll1.Move Width - 380, 0, 250, Height - 650
HScroll1.Move 0, Height - 650, Width - 380, 250
VScroll1.Max = Frame1.Height - Height + 700
HScroll1.Max = Frame1.Width - Width

Corner.Move HScroll1.Width, VScroll1.Height
End Sub

Private Sub Form_Resize()
Scroll_Resize
End Sub

Private Sub Frame1_Click()
Corner.AutoRedraw = True
End Sub

Private Sub HScroll1_Change()
Frame1.Left = -HScroll1.Value
End Sub

Private Sub VScroll1_Change()
Frame1.Top = -VScroll1.Value
End Sub
