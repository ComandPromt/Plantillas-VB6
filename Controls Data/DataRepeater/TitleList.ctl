VERSION 5.00
Begin VB.UserControl TitleList 
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   ScaleHeight     =   585
   ScaleWidth      =   4830
   Begin VB.TextBox txtISBN 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "ISBN"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "TitleList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Title() As String
Attribute Title.VB_MemberFlags = "14"
    Title = txtTitle.Text
End Property

Public Property Let Title(ByVal strNew As String)
    txtTitle.Text = strNew
    PropertyChanged "Title"
End Property

Public Property Get ISBN() As String
Attribute ISBN.VB_MemberFlags = "14"
    ISBN = txtISBN.Text
End Property

Public Property Let ISBN(ByVal strNew As String)
    txtISBN.Text = strNew
    PropertyChanged "ISBN"
End Property
