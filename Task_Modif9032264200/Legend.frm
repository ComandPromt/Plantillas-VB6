VERSION 5.00
Begin VB.Form Legend 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color Info"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2430
   HasDC           =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   Begin VB.Label LLabel 
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape LShape 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Legend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Gray - Not Visible
'Normal - Visible
'Blue - This Application
'Green - (2 total) Computer name & Desktop
'Red - Window Does not exist any longer
'   i could have made it just delete the nodes after finding out the dont exist
'   but i find it better to leave them so that i could see which controls belong to
'   an app if i closeda control
'--------
'other colors
'BOLD & BLACK - found on text boxes throughout the app. these boxes let you change options
'               only when AUTHORMODE is enabled.
'Blue         - When you give focus to a textbox when AUTHORMODE is enabled, the data will be blue
Option Explicit
Private Sub Form_Load()

  Dim i As Integer

    For i = 1 To 4
        Load LShape(i)
        With LShape(i)
            .Move LShape(0).Left, LShape(i - 1).Top + 300
            .Visible = True
        End With 'LSHAPE(I)
        Load LLabel(i)
        With LLabel(i)
            .Move LLabel(0).Left, LLabel(i - 1).Top + 300
            .Visible = True
        End With 'LLABEL(I)
    Next i
    LShape(0).FillColor = vbBlue
    LLabel(0).Caption = "This app"
    LShape(1).FillColor = vbGreen
    LLabel(1).Caption = "Root(Non-Window)"
    LShape(2).FillColor = vbBlack
    LLabel(2).Caption = "Visible window"
    LShape(3).FillColor = RGB(127, 127, 127) 'gray
    LLabel(3).Caption = "Invisible window"
    LShape(4).FillColor = vbRed
    LLabel(4).Caption = "Window no longer exists"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
