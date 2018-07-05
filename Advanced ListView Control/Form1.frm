VERSION 5.00
Object = "{E3A77A65-A5A8-41E0-ABCA-0004D9A78B0B}#6.0#0"; "ADVFIL~1.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin AdvFileList.AdvList AdvList1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _extentx        =   9763
      _extenty        =   7646
      font            =   "Form1.frx":0000
      path            =   "E:\Program Files\Microsoft Visual Studio\VB98\Projects\ActiveX\File_View2"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is an advanced ListView Control. It Populates a listview control using the
'given path with the files and folders. The control retains most listview
'properties, events and methods plus:
'UserControl.Path (Read/Write) Property that specifies the path to populate
'UserControl.Populate Method to populate the control
'Feel free to enhance it!
'Copyright ©  2002 George Kontostanos

Option Explicit
Dim frmHeight As Single, frmWidth As Single

Private Sub Form_Load()
    frmHeight = Form1.Height
    frmWidth = Form1.Width
    AdvList1.Path = "C:\"
    AdvList1.Populate
End Sub

Private Sub Form_Resize()
    AdvList1.Width = AdvList1.Width + (Form1.Width - frmWidth)
    AdvList1.Height = AdvList1.Height + (Form1.Height - frmHeight)
    frmWidth = Form1.Width
    frmHeight = Form1.Height
End Sub
