VERSION 5.00
Begin VB.Form frmPrint 
   Caption         =   "PhoneBook - Print"
   ClientHeight    =   855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2055
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "frmPrint.frx":030A
   ScaleHeight     =   855
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Are You Ready To Print?"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdYes_Click()
    Printer.Font = "arial"
    Printer.FontSize = 18
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print "PhoneBook"
    Printer.Print ""
    Printer.Font = "Arial"
    Printer.FontUnderline = False
    Printer.FontSize = 12
    Printer.Print "Name: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblName.Text
    Printer.FontBold = True
    Printer.Print "Address: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblAddress.Text
    Printer.FontBold = True
    Printer.Print "Suburb: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblSuburb.Text
    Printer.FontBold = True
    Printer.Print "State: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblState.Text
    Printer.FontBold = True
    Printer.Print "Post Code: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblPostCode.Text
    Printer.FontBold = True
    Printer.Print "Country: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblCountry.Text
    Printer.FontBold = True
    Printer.Print "Phone Number: ",
    Printer.FontBold = False
    Printer.Print Form1.lblPhNo.Text
    Printer.FontBold = True
    Printer.Print "Second Phone Number: ",
    Printer.FontBold = False
    Printer.Print Form1.lblPhNo2.Text
    Printer.FontBold = True
    Printer.Print "Fax Number: ",
    Printer.FontBold = False
    Printer.Print Form1.lblFax.Text
    Printer.FontBold = True
    Printer.Print "Mobile Number: ",
    Printer.FontBold = False
    Printer.Print Form1.lblMobile.Text
    Printer.FontBold = True
    Printer.Print "Company Name: ",
    Printer.FontBold = False
    Printer.Print Form1.lblWork.Text
    Printer.FontBold = True
    Printer.Print "Company Ph. No.: ",
    Printer.FontBold = False
    Printer.Print Form1.lblWorkNo.Text
    Printer.FontBold = True
    Printer.Print "Company Fax Number: ",
    Printer.FontBold = False
    Printer.Print Form1.lblCoFax.Text
    Printer.FontBold = True
    Printer.Print "Email: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblEmail.Text
    Printer.FontBold = True
    Printer.Print "Web Site: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblWebSite.Text
    Printer.FontBold = True
    Printer.Print "Comments: ", ,
    Printer.FontBold = False
    Printer.Print Form1.lblComments.Text
    Printer.EndDoc
    frmPrint.Hide
End Sub
Private Sub cmdNo_Click()
    frmPrint.Hide
End Sub
Private Sub Form_Load()
    If Form1.lstNames.ListIndex = -1 Then
        If MsgBox("Warning: You Do Not Have An Entry Selected. If You Continue, Blank Fields Will Be Printed.", vbCritical) = vbOK Then frmPrint.Show
    End If
End Sub
