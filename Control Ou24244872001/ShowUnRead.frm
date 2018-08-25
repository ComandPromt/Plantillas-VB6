VERSION 5.00
Begin VB.Form frmShowUnRead 
   Caption         =   "Form2"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   LinkTopic       =   "Form2"
   ScaleHeight     =   7410
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox txtBody 
      Height          =   2565
      Left            =   1080
      TabIndex        =   3
      Top             =   3840
      Width           =   8535
   End
   Begin VB.TextBox txtRecieved 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   3360
      Width           =   8535
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   2880
      Width           =   8535
   End
   Begin VB.ListBox lstSubjects 
      Height          =   1815
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   8535
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmShowUnRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ol As New Outlook.Application
Private ns As Outlook.NameSpace
Private itms As Outlook.Items

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim fd As Outlook.MAPIFolder
    Dim itm As Object
    Dim l As Long
    
    'Get to the current session of Outlook.
    Set ns = ol.GetNamespace("MAPI")
    
    'Retrieve a collection of mail messages in the inbox.
    Set itms = ns.GetDefaultFolder(olFolderInbox).Items

    'Loop through the items and display the subjects of the unread
    'messages in a list box on a form.
    For Each itm In itms
        If itm.UnRead = True Then
            lstSubjects.AddItem itm.Subject
            l = l + 1
        End If
    Next
    Label1.Caption = CStr(l) & " Un-Read eMail(s) Found"
End Sub

Private Sub lstSubjects_Click()
    Dim itm As Object
    Dim criteria As String

    'Find the mail item with the subject currently selected
    'in a list box called lstSubjects.
    criteria = "[subject] = '" & lstSubjects & "'"
    
    Set itm = itms.Find(criteria)
        
    If itm Is Nothing Then
        MsgBox ("Information not found")
    Else
        
        'Place some item information into text boxes.
        txtFrom.Text = itm.SenderName
        txtRecieved.Text = itm.ReceivedTime
        txtBody.Text = itm.Body
    End If
End Sub
