VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML Invoice Test"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCreateInvoice 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Create HTML Invoice (Uses Outlook)"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Text            =   "email@servername.com.au"
      Top             =   3720
      Width           =   3135
   End
   Begin VB.TextBox txtDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6840
      TabIndex        =   16
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtPurchaseOrder 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6840
      TabIndex        =   15
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtInvoice 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6840
      TabIndex        =   14
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtBillTo 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtShipTo 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "$0.00"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtGST 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "$0.00"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "$0.00"
      Top             =   3600
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":000C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   5040
      Width           =   7935
   End
   Begin VB.Label Label 
      Caption         =   "E-mail Address"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Shape Shape 
      Height          =   1335
      Index           =   1
      Left            =   5400
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label 
      Caption         =   "Date"
      Height          =   255
      Index           =   7
      Left            =   5520
      TabIndex        =   13
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "Purchase Order"
      Height          =   255
      Index           =   6
      Left            =   5520
      TabIndex        =   12
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "Invoice Number"
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "Bill To:"
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   10
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "Ship To:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.Shape Shape 
      Height          =   1215
      Index           =   0
      Left            =   5400
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label 
      Caption         =   "TOTAL AMOUNT"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "GST"
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   5
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "SALE AMOUNT"
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Matt Trigwell 25/02/2003

Private Sub Form_Load()
    
    Call Initialise
    
End Sub


Private Sub Form_Activate()

    txtEmail.SetFocus
    
End Sub

Private Sub Initialise()
    
    'Initialise Grid
    With Grid
        
        'Set Back Colour to forms back colour
        .BackColorBkg = Me.BackColor
        
        'Column Headings
        .TextMatrix(0, 0) = "Quantity"
        .TextMatrix(0, 1) = "B Order"
        .TextMatrix(0, 2) = "Code"
        .TextMatrix(0, 3) = "Description"
        .TextMatrix(0, 4) = "Charges"
        .TextMatrix(0, 5) = "GST"
        .TextMatrix(0, 6) = "Total"
        
        'Column Widths
        .ColWidth(0) = 800
        .ColWidth(1) = 800
        .ColWidth(2) = 800
        .ColWidth(3) = 2700
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        
        'Column Alignments
        .ColAlignment(0) = 4
        .ColAlignment(1) = 4
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        
        'Grid Contents
        .Rows = 4   'Three product rows, One heading row
        
        'First Product
        .TextMatrix(1, 0) = "2"
        .TextMatrix(1, 1) = "0"
        .TextMatrix(1, 2) = "VB6"
        .TextMatrix(1, 3) = "Microsoft Visual Basic 6"
        .TextMatrix(1, 4) = "$1500.00"
        .TextMatrix(1, 5) = "$150.00"
        .TextMatrix(1, 6) = "$1650.00"
        
        'Second Product
        .TextMatrix(2, 0) = "1"
        .TextMatrix(2, 1) = "0"
        .TextMatrix(2, 2) = "SQL2000"
        .TextMatrix(2, 3) = "Microsoft SQL Server 2000"
        .TextMatrix(2, 4) = "$2000.00"
        .TextMatrix(2, 5) = "$200.00"
        .TextMatrix(2, 6) = "$2200.00"
        
        'Third Product
        .TextMatrix(3, 0) = "10"
        .TextMatrix(3, 1) = "4"
        .TextMatrix(3, 2) = "FS2002"
        .TextMatrix(3, 3) = "Flight Simulator 2002"
        .TextMatrix(3, 4) = "$1700.00"
        .TextMatrix(3, 5) = "$170.00"
        .TextMatrix(3, 6) = "$1870.00"

    End With
    
    'Set other data on the form
    txtDate.Text = Date
    txtShipTo.Text = "Industry Computer Programmers" & vbCrLf & "PO Box 2390" & vbCrLf & "Port Adelaide SA  5015" & vbCrLf & "Australia"
    txtBillTo.Text = txtShipTo.Text
    txtInvoice.Text = "873452"
    txtPurchaseOrder.Text = "PO 5329789"
    txtDate.Text = Date
    
    'Totals
    txtAmount.Text = "$5200.00"
    txtGST.Text = "$520.00"
    txtTotal.Text = "$5720.00"
    
    
    'Fill in the Invoice Information - Used when formatting HTML
    InvoiceDetails.InvoiceNumber = txtInvoice.Text
    InvoiceDetails.PurchaseOrder = txtPurchaseOrder.Text
    InvoiceDetails.InvoiceDate = txtDate.Text
    InvoiceDetails.SaleAmount = "$5200.00"
    InvoiceDetails.GST = "$520.00"
    InvoiceDetails.Total = "$5720.00"
    InvoiceDetails.Paid = "$720.00"
    InvoiceDetails.Balance = "$5000.00"
    InvoiceDetails.DueDate = CStr(Date)
    
End Sub


Private Sub cmdCreateInvoice_Click()
    
    On Error GoTo Errorhandler
    
    'Outlook Declares
    Dim objOutlook As Outlook.Application
    Dim objMailItem As Outlook.MailItem
    
    'Create Outlook objects
    Set objOutlook = New Outlook.Application
    Set objMailItem = objOutlook.CreateItem(olMailItem)

    'Display HTML Invoice
    With objMailItem
        .To = txtEmail.Text
        .Subject = "Test HTML Invoice"
        .HTMLBody = FormatHTMLInvoice() 'Create HTML for the invoice
        .Attachments.Add App.Path & "\ftLogo.bmp" 'Attach any pictures required in the HTML document
        .Display
    End With
    
    'Free memory
    Set objMailItem = Nothing
    Set objOutlook = Nothing
    
    Exit Sub
Errorhandler:
    If (Err.Number = 287) Then  'Application Object Error
        Exit Sub
        
    ElseIf (Err.Number = 429) Then  'Outlook not installed
        MsgBox "Microsoft Outlook is required to access this function.", vbInformation, "Resource Not Available"
    
    Else    'Other error
        MsgBox Err.Description, vbExclamation, Err.Number
        Debug.Print Err.Number
    End If

    
End Sub


Private Function InvoiceItems() As String

    'Returns formatted HTML
    'Adds the items to the invoice
    
    Dim Index As Integer
    Dim NumberItems As Integer
    Dim strReturn As String
    
    With Grid
        
        'Count number of items
        NumberItems = .Rows - 1
        
        For Index = 1 To NumberItems
        
            strReturn = strReturn & _
               "<tr>" & _
               "<td width=""8%"" align=""center""><font face=""Arial"" size=""2"">" & .TextMatrix(Index, 0) & "</font></td>" & _
               "<td width=""8%"" align=""center""><font face=""Arial"" size=""2"">" & .TextMatrix(Index, 1) & "</font></td>" & _
               "<td width=""8%"" align=""left""><font face=""Arial"" size=""2"">" & .TextMatrix(Index, 2) & "</font></td>" & _
               "<td width=""25%"" align=""left""><font face=""Arial"" size=""2"">" & .TextMatrix(Index, 3) & "</font></td>" & _
               "<td width=""10%"" align=""right""><font face=""Arial"" size=""2"">" & .TextMatrix(Index, 4) & "</font></td>" & _
               "<td width=""10%"" align=""right""><font face=""Arial"" size=""2"">" & .TextMatrix(Index, 5) & "</font></td>" & _
               "<td width=""10%"" align=""right""><font face=""Arial"" size=""2"">" & .TextMatrix(Index, 6) & "</font></td>" & _
               "</tr>"
               
        Next
        
    End With
    
    'Return
    InvoiceItems = strReturn
  
       
End Function

Private Function AddressInfo(AddressText As String) As String
    
    'Returns formatted HTML
    'Formats multiline address information
    
    Dim strHTMLAddress As String
    Dim Position As Integer
    
    'Change vbCrLf to <br>
    strHTMLAddress = Replace(AddressText, vbCrLf, "<br>")
    
    'Bold first line
    Position = InStr(1, strHTMLAddress, "<br>")
    If (Position > 0) Then
        strHTMLAddress = "<b>" & Left(strHTMLAddress, Position - 1) & "</b>" & Right(strHTMLAddress, Len(strHTMLAddress) - (Position - 1))
    End If
    
    AddressInfo = strHTMLAddress
    
End Function

Private Function FormatHTMLInvoice() As String
    
    'Function returns a string containing a html code for an invoice based on the
    'data entered into the form.
    
    Dim strHTMLInvoice As String
    
    'Start of HTML
    strHTMLInvoice = "<html><head>" & _
       "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">" & "<title>Invoice</title>" & "</head>"
                              
    'Company Information
    strHTMLInvoice = strHTMLInvoice & _
       "<body> " & "<table border=""0"" cellpadding=""4"" & cellspacing=""5"" style=""border-collapse: collapse"" bordercolor=""#000000"" width=""725""><tr>" & "<td width=""583""> " & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#FFFFFF"" width=""700"" height=""538""><tr> " & _
       "<td width=""100%"" height=""158""> " & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#FFFFFF"" width=""100%"" bgcolor=""#FFFFFF"" height=""158""><tr> " & "<td width=""17%"" align=""left"" valign=""top"" height=""158""> " & "<img border=""0"" src=""ftLogo.bmp"" hspace=""5"" vspace=""5"" width=""273"" height=""123""></td> " & _
       "<td width=""40%"" align=""left"" valign=""top"" height=""158""> " & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" width=""74%""  height=""67"" bordercolor=""#FFFFFF""><tr> " & "<td width=""100%"" height=""39""><font face=""Arial"" size=""1""><br>Force-techie Limited<br>Storage Computer House<br>Cleeve Road<br>Leatherhead<br>Surrey<br>KT22 7NB<br>&nbsp;</font></td></tr><tr> " & "<td width=""100%"" height=""27""><font size=""1"" face=""Arial"">Phone: 01372 374758<br>Fax: 01372 361853<br><a href=""http://www.planetsourcecode.com"">www.planetsourcecode.com</a></font><br>&nbsp;</td></tr></table></td> " & "<td width=""100%"" align=""right"" valign=""top"" height=""158"">&nbsp;"
       
    'Invoice Number, Date etc..
    strHTMLInvoice = strHTMLInvoice & _
       "<table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""border-collapse: collapse"" bordercolor=""#000000"" width=""240""><tr> " & _
        "<td width=""100%"" bgcolor=""#D6D6D6""><p align=""center""><b><font face=""Arial"" size=""4"">TAX&nbsp; INVOICE</font></b></td></tr><tr> " & "<td width=""100%"">" & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#FFFFFF"" width=""91%""><tr>" & "<td width=""50%""><font face=""Arial"" size=""2"">&nbsp;INV#</font></td> " & "<td width=""50%""><font face=""Arial"" size=""2"">" & InvoiceDetails.InvoiceNumber & "</font></td></tr><tr> " & "<td width=""50%""><font face=""Arial"" size=""2"">&nbsp;P Order #</font></td> " & "<td width=""50%""><font face=""Arial"" size=""2"">" & InvoiceDetails.PurchaseOrder & "</font></td></tr><tr> " & "<td width=""50%""><font face=""Arial"" size=""2"">&nbsp;DATE</font></td> " & "<td width=""50%""><font face=""Arial"" size=""2"">" & InvoiceDetails.InvoiceDate & "</font></td> " & "</tr></table></td></tr></table></td><td width=""65%"" align=""right"" height=""158""><br><br><br>&nbsp;</td></tr></table></td></tr><tr>"
  
    'Address Information Start
    strHTMLInvoice = strHTMLInvoice & "<td width=""100%"" height=""104""> " & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#FFFFFF"" width=""100%""><tr> " & "<td width=""43%"">"
       
    'Address Ship To
    strHTMLInvoice = strHTMLInvoice & _
        "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#FFFFFF"" width=""100%""><tr> " & "<td width=""100%""><font face=""Arial"" size=""1"">&nbsp; Ship To:<br>&nbsp;</font></td></tr><tr> " & "<td width=""100%""><blockquote><p><font face=""Arial"" size=""2"">" & AddressInfo(txtShipTo.Text) & "</font></p></blockquote></td></tr></table></td> " & "<td width=""14%"">&nbsp;</td> " & "<td width=""40%"" align=""left"" valign=""top"">"

    'Address Ship To
    strHTMLInvoice = strHTMLInvoice & _
        "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#FFFFFF"" width=""200%""><tr> " & "<td width=""100%""><font face=""Arial"" size=""1"">Bill To:<br>&nbsp;</font></td></tr><tr> " & "<td width=""100%""><blockquote><p><font face=""Arial"" size=""2"">" & AddressInfo(txtBillTo.Text) & "</font></p></blockquote></td></tr></table></td> " & "<td width=""3%"">&nbsp;</td></tr></table></td></tr><tr> " & "<td width=""100%"" height=""34""><font face=""Arial"" size=""2"">&nbsp; </font>"
    
    'Invoice column headings
    strHTMLInvoice = strHTMLInvoice & _
        "<table border=""1"" cellpadding=""1"" cellspacing=""1"" style=""border-collapse: collapse"" bordercolor=""#000000"" width=""100%""><tr> " & "<td width=""8%"" bgcolor=""#D6D6D6"" align=""center""><b><font face=""Arial"" size=""2"">Quantity</font></b></td> " & "<td width=""8%"" bgcolor=""#D6D6D6"" align=""center""><b><font face=""Arial"" size=""2"">B Order</font></b></td> " & "<td width=""8%"" bgcolor=""#D6D6D6"" align=""center""><b><font face=""Arial"" size=""2"">Code</font></b></td> " & "<td width=""25%"" bgcolor=""#D6D6D6"" align=""center""><b><font face=""Arial"" size=""2"">Description</font></b></td> " & "<td width=""10%"" bgcolor=""#D6D6D6"" align=""center""><b><font face=""Arial"" size=""2"">Charges</font></b></td> " & "<td width=""10%"" bgcolor=""#D6D6D6"" align=""center""><b><font face=""Arial"" size=""2"">GST</font></b></td> " & "<td width=""10%"" bgcolor=""#D6D6D6"" align=""center""><b><font face=""Arial"" size=""2"">Total</font></b></td> " & "</tr></table></td></tr>"
    
    'Invoice Items Start
    strHTMLInvoice = strHTMLInvoice & _
       "<tr>" & "<td width=""100%"" height=""19"">" & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#111111"" width=""100%"" id=""AutoNumber1"">"
    'Invoice Items Contents
    strHTMLInvoice = strHTMLInvoice & InvoiceItems()
       
    'Invoice Items End
    strHTMLInvoice = strHTMLInvoice & _
       "</table>" & "</td>" & "</tr>"

    'Notice
    strHTMLInvoice = strHTMLInvoice & _
       "<tr><td width=""100%"" height=""90""><table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#FFFFFF"" width=""100%""><tr>" & _
        "<td width=""100%""><BR>" & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#FFFFFF"" width=""100%""><tr>" & "<td width=""18%""><b><font face=""Arial"" size=""4"">&nbsp;Please Note:</font></b></td> " & "<td width=""82%"" bgcolor=""#FFFFFF""><b><font face=""Arial"" size=""2"">&nbsp;Transport Costs are the responsibility of the purchaser.<br>&nbsp;This includes those goods returning for after sales service.<br> &nbsp;American Express cards incur a 4% surcharge.&nbsp; Other credit cards incur 2%</font></b></td></tr></table></td></tr></table></td></tr><tr>" & "<td width=""100%"" height=""128""><br>&nbsp;<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#FFFFFF"" width=""100%""><tr>" & "<td width=""29%"" align=""left"" valign=""top"">"
    
    
    'Totals Part 1
    strHTMLInvoice = strHTMLInvoice & _
       "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#FFFFFF"" width=""100%""><tr>" & _
       "<td width=""53%""><font face=""Arial"" size=""2"">&nbsp;<b>Payment Due</b></font></td> " & _
       "<td width=""47%"" align=""center"">&nbsp;<font face=""Arial"" size=""2"">" & InvoiceDetails.DueDate & "</font></td></tr></table></td> " & _
       "<td width=""37%"">&nbsp;</td> " & _
       "<td width=""34%""> " & _
       "<table border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#111111"" width=""100%""><tr>" & _
       "<td width=""50%""> " & _
       "<table border=""0"" cellpadding=""1"" cellspacing=""1"" style=""border-collapse: collapse"" bordercolor=""#FFFFFF"" width=""100%"" align=""right""><tr>" & _
       "<td width=""100%"" align=""right""><font face=""Arial"" size=""2"">SALE AMOUNT&nbsp;&nbsp;</font></td></tr><tr>" & _
       "<td width=""100%"" align=""right""><font face=""Arial"" size=""2"">GST&nbsp;&nbsp;</font></td></tr><tr> " & _
       "<td width=""100%"" align=""right""><font face=""Arial"" size=""2"">TOTAL&nbsp;&nbsp;</font></td></tr><tr>" & _
       "<td width=""100%"" align=""right""><font face=""Arial"" size=""2"">PAID&nbsp;&nbsp;</font></td></tr><tr> " & _
       "<td width=""100%"" align=""right""><font face=""Arial"" size=""2""><b>BALANCE&nbsp;&nbsp;</b></font></td></tr></table></td>" & _
       "<td width=""50%"" align=""right"">"
    
    'Totals Part 2
    strHTMLInvoice = strHTMLInvoice & _
       "<table border=""0"" cellpadding=""1"" cellspacing=""1"" style=""border-collapse: collapse"" bordercolor=""#FFFFFF"" width=""100%"" align=""right""><tr>" & _
       "<td width=""100%"" bordercolor=""#FFFFFF"" bgcolor=""#FFFFFF"" align=""right""><b><font face=""Arial"" size=""2"">" & InvoiceDetails.SaleAmount & "</font></b></td></tr><tr>" & _
       "<td width=""100%"" bordercolor=""#FFFFFF"" bgcolor=""#FFFFFF"" align=""right""><b><font face=""Arial"" size=""2"">" & InvoiceDetails.GST & "</font></b></td></tr><tr>" & _
       "<td width=""100%"" bordercolor=""#FFFFFF"" bgcolor=""#FFFFFF"" align=""right""><b><font face=""Arial"" size=""2"">" & InvoiceDetails.Total & "</font></b></td></tr><tr>" & _
       "<td width=""100%"" bordercolor=""#FFFFFF"" bgcolor=""#FFFFFF"" align=""right""><b><font face=""Arial"" size=""2"">" & InvoiceDetails.Paid & "</font></b></td></tr><tr>" & _
       "<td width=""100%"" bordercolor=""#FFFFFF"" bgcolor=""#D6D6D6"" align=""right""><b><font face=""Arial"" size=""2"">" & InvoiceDetails.Balance & "</font></b></td></tr>" & _
       "</table></td></tr></table></td></tr></table></td></tr></table></td></tr></table>" & _
       "</body></html>"
       

    'Return formatted html string
    FormatHTMLInvoice = strHTMLInvoice

End Function


