VERSION 5.00
Object = "{D28F8786-0BB9-402B-92DC-F32DE23A324E}#3.0#0"; "OutlookBar.ocx"
Begin VB.Form Form2 
   Caption         =   "More Samples"
   ClientHeight    =   5388
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   6180
   LinkTopic       =   "Form2"
   ScaleHeight     =   5388
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin OutlookBar.ctxOutlookBar ctxOutlookBar1 
      Height          =   5052
      Left            =   84
      TabIndex        =   0
      Top             =   168
      Width           =   1524
      _ExtentX        =   2709
      _ExtentY        =   8911
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormatControl   =   "Form2.frx":0000
      FormatGroup     =   "Form2.frx":0174
      FormatGroupHover=   "Form2.frx":0234
      FormatGroupPressed=   "Form2.frx":02F4
      FormatGroupSelected=   "Form2.frx":03C8
      FormatItem      =   "Form2.frx":0474
      FormatItemLargeIcons=   "Form2.frx":0548
      FormatItemHover =   "Form2.frx":0630
      FormatItemPressed=   "Form2.frx":06DC
      FormatItemSelected=   "Form2.frx":0788
      FormatSmallIcon =   "Form2.frx":0834
      FormatSmallIconHover=   "Form2.frx":091C
      FormatSmallIconPressed=   "Form2.frx":0A18
      FormatSmallIconSelected=   "Form2.frx":0B14
      FormatLargeIcon =   "Form2.frx":0C10
      FormatLargeIconHover=   "Form2.frx":0CF8
      FormatLargeIconPressed=   "Form2.frx":0DF4
      FormatLargeIconSelected=   "Form2.frx":0EF0
      Groups          =   "Form2.frx":0FEC
      LabelEdit       =   0
   End
   Begin OutlookBar.ctxOutlookBar ctxOutlookBar2 
      Height          =   5052
      Left            =   1848
      TabIndex        =   1
      Top             =   168
      Width           =   1524
      _ExtentX        =   2709
      _ExtentY        =   8911
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormatControl   =   "Form2.frx":24A4
      FormatGroup     =   "Form2.frx":25F0
      FormatGroupHover=   "Form2.frx":26B0
      FormatGroupPressed=   "Form2.frx":2770
      FormatGroupSelected=   "Form2.frx":2844
      FormatItem      =   "Form2.frx":28F0
      FormatItemLargeIcons=   "Form2.frx":29D8
      FormatItemHover =   "Form2.frx":2AD4
      FormatItemPressed=   "Form2.frx":2B94
      FormatItemSelected=   "Form2.frx":2C54
      FormatSmallIcon =   "Form2.frx":2D14
      FormatSmallIconHover=   "Form2.frx":2DFC
      FormatSmallIconPressed=   "Form2.frx":2ED0
      FormatSmallIconSelected=   "Form2.frx":2FB8
      FormatLargeIcon =   "Form2.frx":308C
      FormatLargeIconHover=   "Form2.frx":3188
      FormatLargeIconPressed=   "Form2.frx":325C
      FormatLargeIconSelected=   "Form2.frx":3344
      Groups          =   "Form2.frx":3418
      FlatScrollArrows=   0   'False
      LabelEdit       =   0
   End
   Begin OutlookBar.ctxOutlookBar ctxOutlookBar3 
      Height          =   5052
      Left            =   3612
      TabIndex        =   2
      Top             =   168
      Width           =   1524
      _ExtentX        =   2709
      _ExtentY        =   8911
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormatControl   =   "Form2.frx":492C
      FormatGroup     =   "Form2.frx":4A8C
      FormatGroupHover=   "Form2.frx":4B74
      FormatGroupPressed=   "Form2.frx":4C34
      FormatGroupSelected=   "Form2.frx":4D1C
      FormatItem      =   "Form2.frx":4DC8
      FormatItemLargeIcons=   "Form2.frx":4EC4
      FormatItemHover =   "Form2.frx":4FD4
      FormatItemPressed=   "Form2.frx":5100
      FormatItemSelected=   "Form2.frx":5240
      FormatSmallIcon =   "Form2.frx":536C
      FormatSmallIconHover=   "Form2.frx":5454
      FormatSmallIconPressed=   "Form2.frx":553C
      FormatSmallIconSelected=   "Form2.frx":5638
      FormatLargeIcon =   "Form2.frx":5720
      FormatLargeIconHover=   "Form2.frx":581C
      FormatLargeIconPressed=   "Form2.frx":5904
      FormatLargeIconSelected=   "Form2.frx":5A00
      Groups          =   "Form2.frx":5AE8
      FlatScrollArrows=   0   'False
      LabelEdit       =   0
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ctxOutlookBar1_ButtonClick(ByVal oBtn As OutlookBar.cButton)
    If oBtn.Index = 1 Then
        ctxOutlookBar1.Groups(1).IconsType = 1 - ctxOutlookBar1.Groups(1).IconsType
    End If
End Sub

Private Sub ctxOutlookBar2_ButtonClick(ByVal oBtn As OutlookBar.cButton)
    If oBtn.Index = 1 Then
        ctxOutlookBar2.Groups(1).IconsType = 1 - ctxOutlookBar2.Groups(1).IconsType
    End If
End Sub

Private Sub ctxOutlookBar3_ButtonClick(ByVal oBtn As OutlookBar.cButton)
    If oBtn.Index = 1 Then
        ctxOutlookBar3.Groups(1).IconsType = 1 - ctxOutlookBar3.Groups(1).IconsType
    End If
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    ctxOutlookBar1.Height = ScaleHeight - 2 * ctxOutlookBar1.Top
    ctxOutlookBar2.Height = ScaleHeight - 2 * ctxOutlookBar2.Top
    ctxOutlookBar3.Height = ScaleHeight - 2 * ctxOutlookBar3.Top
End Sub
