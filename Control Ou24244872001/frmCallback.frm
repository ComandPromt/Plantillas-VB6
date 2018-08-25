VERSION 5.00
Begin VB.Form frmCallback 
   Caption         =   "frmCallback"
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "Try to send mail now."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents objApplication As Outlook.Application
Attribute objApplication.VB_VarHelpID = -1
'Dim ol As New Outlook.Application
Dim ns As Outlook.NameSpace
Private Sub Form_Load()
    'then in the startup code
    Set objApplication = Application
       'Return a reference to the MAPI layer.
    Set ns = objApplication.GetNamespace("MAPI")
    ns.Logon
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objApplication = Nothing
    Set ns = Nothing
    
End Sub

'then the event code
Sub objApplication_ItemSend(ByVal Item As Object, Cancel As Boolean)

    MsgBox "Sending mail is blocked by this program"
    Cancel = True

End Sub

'then the event code
Sub objApplication_NewMail()

    MsgBox "New mail has arrrived"

End Sub
