VERSION 5.00
Begin VB.Form formADOXML 
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "FromXML"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ToXML"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "ADOXML.frx":0000
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "formADOXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objRS   As ADODB.Recordset

Private Sub Command1_Click(Index As Integer)
    Dim objStream As Stream
    
    Set objStream = New Stream
    
    Select Case Index
    Case 0 'ToXML
        objRS.Save objStream, adPersistXML
        Text1.Text = objStream.ReadText
    
    Case 1 'FromXML
        objStream.Open
        objStream.WriteText Text1.Text
        objStream.Position = 0
        
        objRS.Close
        objRS.Open objStream
    End Select
    
    objStream.Close
    Set objStream = Nothing

End Sub

Private Sub Form_Load()
    Dim intcnt As Integer
    'Create Set
    Set objRS = New ADODB.Recordset
    objRS.Fields.Append "NAME", adVarChar, 255
    objRS.Fields.Append "VALUE", adVariant
    objRS.Open
    
    'Add Records
    For intcnt = 1 To 5
        objRS.AddNew
        objRS.Fields("NAME").Value = "ITEM " & intcnt
        objRS.Fields("VALUE").Value = "VALUE " & intcnt
        objRS.Update
    Next
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Text1.Width = ScaleWidth
    Text1.Height = ScaleHeight - Text1.Top
End Sub
