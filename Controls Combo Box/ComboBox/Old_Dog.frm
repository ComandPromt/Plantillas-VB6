VERSION 5.00
Begin VB.Form Old_Dog 
   Caption         =   "Old Dog"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Add New Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Old_Dog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hWnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, _
lParam As Any) As Long

Private Declare Function SendMessageByString Lib "user32" _
Alias "SendMessageA" (ByVal hWnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, _
lParam As String) As Long

Const CB_SHOWDROPDOWN = &H14F
Const CB_FINDSTRINGEXACT = &H158

Private Sub Command1_Click()
    With Combo1
        If Not CheckDuplicates(.hWnd, .Text) Then
            .AddItem .Text
        End If
        SendMessage .hWnd, CB_SHOWDROPDOWN, _
            True, ByVal 0&
    End With
    
End Sub

Private Sub Form_Load()
    'Put some items in the combobox
    With Combo1
        .AddItem "Sunday"
        .AddItem "Monday"
        .AddItem "Tuesday"
        .AddItem "Wednesday"
        .AddItem "Thursday"
        .AddItem "Friday"
        .AddItem "Saturday"
        
        'Provide the default value
        .Text = .List(0)
        
        'Open the dropdown list
        SendMessage .hWnd, CB_SHOWDROPDOWN, _
            True, ByVal 0&
        
    End With
End Sub
Private Function CheckDuplicates(chwnd As Long, _
    StringText As String) As Boolean

CheckDuplicates = SendMessageByString(chwnd, CB_FINDSTRINGEXACT, _
    -1, ByVal StringText) > -1

End Function

