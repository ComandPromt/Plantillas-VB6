VERSION 5.00
Begin VB.Form Dialog1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dialog Caption"
   ClientHeight    =   2400
   ClientLeft      =   3900
   ClientTop       =   3720
   ClientWidth     =   4890
   Icon            =   "Time.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Dialog1.Hide
End Sub

Private Sub Form_Load()
List1.AddItem Format(Now, "long time")
List1.AddItem Format(Now, "short time")
List1.AddItem Format(Now, "medium time")
List1.AddItem Format(Now, "general date")
List1.AddItem Format(Now, "long date")
List1.AddItem Format(Now, "medium date")
List1.AddItem Format(Now, "short date")
List1.AddItem (Date)
List1.AddItem Format(Date, "dd - mm - yyyy")
List1.AddItem Format(Date, "dd-mm-yy")
List1.AddItem Format(Date, "dd/mm/yy")
List1.AddItem Format(Date, "dd/mm/yyyy")
List1.AddItem Format(Date, "dd/mm")
List1.AddItem Format(Date, "dd")
List1.AddItem Format(Time, "hh-mm-ss")
List1.AddItem Format(Time, "hh.mm.ss")
List1.AddItem Format(Time, "hh-mm")
End Sub


Private Sub List1_DblClick()
    OKButton_Click
End Sub

Private Sub OKButton_Click()
    fMainForm.Text1.SelText = List1.Text
    Unload Me
End Sub
