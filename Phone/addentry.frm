VERSION 5.00
Begin VB.Form frmAddEntry 
   Caption         =   "Add New PhoneBook Entry"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5235
   Icon            =   "addentry.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "addentry.frx":030A
   ScaleHeight     =   6300
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPostCode 
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtCountry 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox txtState 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtSuburb 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   3855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Entry"
      DownPicture     =   "addentry.frx":1A78C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtComments 
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Top             =   5760
      Width           =   3855
   End
   Begin VB.TextBox txtPhNo 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox txtWebSite 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   5400
      Width           =   3855
   End
   Begin VB.TextBox txtCoFax 
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Top             =   4680
      Width           =   3855
   End
   Begin VB.TextBox txtWorkNo 
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Top             =   4320
      Width           =   3855
   End
   Begin VB.TextBox txtWork 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   3960
      Width           =   3855
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Top             =   5040
      Width           =   3855
   End
   Begin VB.TextBox txtMobile 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox txtFax 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   3240
      Width           =   3855
   End
   Begin VB.TextBox txtPhNo2 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   2880
      Width           =   3855
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Suburb:"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code:"
      Height          =   255
      Left            =   2760
      TabIndex        =   31
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Web Site:"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Fax:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "2nd Ph. Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblComments 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblWorkNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Ph No:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblWork 
      BackStyle       =   0  'Transparent
      Caption         =   "Company:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lblMobile 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblPhNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblAddEntry 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add New PhoneBook Entry."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmAddEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    If txtName = "" Then
        MsgBox "You Must Type In A Name To Add", vbExclamation, "PhoneBook"
    Else
        Form1.lstNames.AddItem txtName.Text
        Form1.lstNumbers.AddItem txtPhNo.Text
        Form1.lstNumbers2.AddItem txtPhNo2.Text
        Form1.lstFax.AddItem txtFax.Text
        Form1.lstMobile.AddItem txtMobile.Text
        Form1.lstEmail.AddItem txtEmail.Text
        Form1.lstWork.AddItem txtWork.Text
        Form1.lstWorkNo.AddItem txtWorkNo.Text
        Form1.lstCoFax.AddItem txtCoFax.Text
        Form1.lstWebSite.AddItem txtWebSite.Text
        Form1.lstaddress.AddItem txtAddress.Text
        Form1.lstComments.AddItem txtComments.Text
        Form1.lstSuburb.AddItem txtSuburb.Text
        Form1.lstState.AddItem txtState.Text
        Form1.lstPostCode.AddItem txtPostCode.Text
        Form1.lstCountry.AddItem txtCountry.Text
        txtName.Text = ""
        txtPhNo.Text = ""
        txtPhNo2.Text = ""
        txtFax.Text = ""
        txtMobile.Text = ""
        txtEmail.Text = ""
        txtWork.Text = ""
        txtWorkNo.Text = ""
        txtCoFax.Text = ""
        txtWebSite.Text = ""
        txtAddress.Text = ""
        txtComments.Text = ""
        txtSuburb.Text = ""
        txtState.Text = ""
        txtPostCode.Text = ""
        txtCountry.Text = ""
        Open "Numbers.dat" For Output As 1
        For i = 0 To Form1.lstNames.ListCount - 1
        Print #1, Form1.lstNames.List(i)
        Print #1, Form1.lstaddress.List(i)
        Print #1, Form1.lstSuburb.List(i)
        Print #1, Form1.lstState.List(i)
        Print #1, Form1.lstPostCode.List(i)
        Print #1, Form1.lstCountry.List(i)
        Print #1, Form1.lstNumbers.List(i)
        Print #1, Form1.lstNumbers2.List(i)
        Print #1, Form1.lstFax.List(i)
        Print #1, Form1.lstMobile.List(i)
        Print #1, Form1.lstWork.List(i)
        Print #1, Form1.lstWorkNo.List(i)
        Print #1, Form1.lstCoFax.List(i)
        Print #1, Form1.lstEmail.List(i)
        Print #1, Form1.lstWebSite.List(i)
        Print #1, Form1.lstComments.List(i)
        Next i
    Close #1
        MsgBox "Your Entry Has Been Added.", vbInformation, "PhoneBook"
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtSuburb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtState_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtPostCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtCountry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtPhNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtPhNo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtMobile_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtWork_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtWorkNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtCoFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtWebSite_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtComments_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAdd_Click
    End If
End Sub
Private Sub txtState_LostFocus()
    txtState.Text = UCase(txtState.Text)
End Sub

