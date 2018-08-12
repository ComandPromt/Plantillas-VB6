VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "How to use the RES Protocol in VB"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Show page..."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6615
      ExtentX         =   11668
      ExtentY         =   8916
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Now let us navigate through our resource file
    
    
    WebBrowser1.Navigate "res://" & App.Path & "\" & App.EXEName & ".exe/page.htm"
    
End Sub

