VERSION 5.00
Begin VB.Form frmWSClient 
   Caption         =   "Web Services Client"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtXML 
      Height          =   6615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1440
      Width           =   12375
   End
   Begin VB.TextBox txtParm1 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Text            =   "ALFKI"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Height          =   375
      Left            =   10320
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtWSDL 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Text            =   "http://etier3/Northwind/soapprocedures?wsdl"
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   "Parm1:"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "WSDL File:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmWSClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oSoapclient As SoapClient


Private Sub cmdExecute_Click()
    Dim oReturnNodeList As IXMLDOMNodeList
    Dim oNode As IXMLDOMNode
    Dim vRetValue As Variant
    Dim sXMLOutput As String
    Dim sEndPointURL As String
    
    Set oSoapclient = New SoapClient
    
    Call oSoapclient.mssoapinit(txtWSDL.Text)
    sEndPointURL = Mid(txtWSDL.Text, 1, (Len(txtWSDL.Text) - 5))
    oSoapclient.ConnectorProperty("EndPointURL") = sEndPointURL
        
    Set oReturnNodeList = oSoapclient.CustOrderHist(txtParm1.Text)
    sXMLOutput = "CustOrderHist Storeprocedure -------------------"
    For Each oNode In oReturnNodeList
        sXMLOutput = sXMLOutput & oNode.xml
    Next
    
    sXMLOutput = sXMLOutput & "GetAllCustomers Template -----------------"
    Set oReturnNodeList = oSoapclient.GetAllCustomers()
    For Each oNode In oReturnNodeList
        sXMLOutput = sXMLOutput & oNode.xml
    Next
    
    sXMLOutput = sXMLOutput & "GetCustomerContactView UDF -----------------"
    Set oReturnNodeList = oSoapclient.GetCustomerContactView(txtParm1.Text)
    For Each oNode In oReturnNodeList
        sXMLOutput = sXMLOutput & oNode.xml
    Next
        
    txtXML.Text = sXMLOutput
    Set oSoapclient = Nothing
End Sub

