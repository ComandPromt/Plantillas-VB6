VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmInvoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Generator"
   ClientHeight    =   10530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11460
   Icon            =   "frmInvoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   255
      Left            =   10320
      TabIndex        =   3
      Top             =   10200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print Two"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   10200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print One"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   10200
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   17806
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmInvoice.frx":000C
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.SelPrint Printer.hDC
End Sub

Private Sub Command2_Click()
    Text1.SelPrint Printer.hDC
    Text1.SelPrint Printer.hDC
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Temp As String
    Dim i As Integer
    Dim Data As String
    Dim Temp2 As String
    Dim Header As String
    Dim Total As String
    Dim Op As Double
    Dim CT As Integer
    Dim r As Integer
    
    Total = 0
    
    Text1.LoadFile App.Path & "\DefaultInvoice.rtf"
    DoEvents
    'Temp = Text1.Text
    Temp = Text1.TextRTF
    
    Temp = Replace(Temp, "Their Address Name", Space(9) & frmMain.txtBillingName.Text)
    Temp = Replace(Temp, "Their Address Line1", Space(9) & frmMain.txtBillingAddress.Text)
    Temp = Replace(Temp, "Their Address Line2", Space(9) & frmMain.txtCityStateZip.Text)
    Temp = Replace(Temp, "Their Address Line3", Space(9) & frmMain.txtOther.Text)

    Temp = Replace(Temp, "0000", Format(frmMain.txtFields(2).Text, "####")) 'Set the invoice number
    Temp = Replace(Temp, "xx/xx/xx", (Format(Date, "mm")) & "/01/" & Format(Date, "yy"))  'Set the invoice date
    
    Temp = Replace(Temp, "Month", frmMain.txtPC.Text)  'Set the invoice date
    
    For i = 1 To frmMain.txtFields(3).Text - 1
        Temp = Replace(Temp, "\par \plain\f3\fs24 Lines", ("\par \plain\f3\fs24 Lines" & vbCrLf & "\par \plain\f3\fs24 Lines"), 1, 1)
        Temp = Replace(Temp, "\par \plain\f2\fs24 Lines", ("\par \plain\f2\fs24 Lines" & vbCrLf & "\par \plain\f2\fs24 Lines"), 1, 1)
    Next i
    
    Do Until frmMain.Op = 1
        frmMain.Command8_Click
    Loop
    
    Header = ""
    
    If frmMain.txtOperations.Text = "" Then frmMain.txtOperations.Text = 1
    For i = 1 To frmMain.txtFields(3).Text
        Data = ""
        CT = 1
        If frmMain.txtDate.Enabled = True Then
            Data = frmMain.txtDate.Text
            If i = 1 Then Header = Header & "\tab Date"
            CT = CT + 1
        End If
        If frmMain.txtAirbill.Enabled = True Then
            If CT = 1 Then
                Data = Data & frmMain.txtAirbill.Text
            Else
                Data = Data & "\tab " & frmMain.txtAirbill.Text
            End If
            CT = CT + 1
            If i = 1 Then Header = Header & "\tab Airbill"
        End If
        If frmMain.txtAircraft.Enabled = True Then
            If CT = 1 Then
                Data = Data & frmMain.txtAircraft.Text
            Else
                Data = Data & "\tab " & frmMain.txtAircraft.Text
            End If
            CT = CT + 1
            If i = 1 Then Header = Header & "\tab Aircraft #"
        End If
        If frmMain.txtDescription.Enabled = True Then
            If CT = 1 Then
                Data = Data & frmMain.txtDescription.Text
            Else
                Data = Data & "\tab " & frmMain.txtDescription.Text
            End If
            CT = CT + 1
            If i = 1 Then Header = Header & "\tab Description"
        End If
        If frmMain.txtDestination.Enabled = True Then
            If CT = 1 Then
                Data = Data & frmMain.txtDestination.Text
            Else
                Data = Data & "\tab " & frmMain.txtDestination.Text
            End If
            CT = CT + 1
            If i = 1 Then Header = Header & "\tab Destination"
        End If
        If frmMain.txtPO.Enabled = True Then
            If CT = 1 Then
                Data = Data & frmMain.txtPO.Text
            Else
                Data = Data & "\tab " & frmMain.txtPO.Text
            End If
            CT = CT + 1
            If i = 1 Then Header = Header & "\tab Purchase Order #"
        End If
        If frmMain.txtFlight.Enabled = True Then
            If CT = 1 Then
                Data = Data & frmMain.txtFlight.Text
            Else
                Data = Data & "\tab " & frmMain.txtFlight.Text
            End If
            CT = CT + 1
            If i = 1 Then Header = Header & "\tab Flight #"
        End If
        If frmMain.txtWeight.Enabled = True Then
            If Left(frmMain.txtWeight.Text, 2) = "1 " Then GoTo SkipOp2
            If CT = 1 Then
                Data = Data & frmMain.txtWeight.Text
            Else
                Data = Data & "\tab " & frmMain.txtWeight.Text
            End If
            CT = CT + 1
            If i = 1 Then Header = Header & "\tab Weight (lbs.)"
            Temp2 = frmMain.txtWeight.Text
            frmMain.txtWeight.Text = ""
            For r = 1 To Len(Temp2) + 1
                If IsNumeric(Mid(Temp2, r, 1)) Then
                    frmMain.txtWeight.Text = frmMain.txtWeight.Text & Mid(Temp2, r, 1)
                End If
            Next r
            Op = Op + CDbl(frmMain.txtWeight.Text)
            frmMain.txtWeight.Text = Temp2
SkipOp2:
        End If
        If frmMain.txtOperations.Enabled = True Then
            If frmMain.txtOperations.Text = 1 And frmMain.txtFields(3).Text <= 1 Then Op = 1: GoTo SkipOp
            If frmMain.txtOperations.Text = 1 Then Op = 0
            If CT = 1 Then
                Data = Data & frmMain.txtOperations.Text
            Else
                Data = Data & "\tab " & frmMain.txtOperations.Text
            End If
            CT = CT + 1
            If i = 1 Then Header = Header & "\tab Operations"
            Op = Op + CDbl(frmMain.txtOperations.Text)
SkipOp:
        End If
'        If i = 1 Then
'            Data = Right(Data, Len(Data) - 5)
'        End If
        
            If Op = 0 Then Op = 1
            Data = Data & "\tab " & frmMain.txtFields(5).Text & "\tab " & Format(frmMain.txtFields(5).Text * frmMain.txtOperations.Text, "$#0.00")
            Total = Total + (frmMain.txtFields(5).Text * frmMain.txtOperations.Text)
    
        Temp = Replace(Temp, "Lines", Data, , 1)
        frmMain.Command7_Click
        If frmMain.txtOperations.Text = "" Then frmMain.txtOperations.Text = 1
    Next i
    
    If frmMain.txtOperations.Text = "1" Then frmMain.txtOperations.Text = ""
    Do Until frmMain.Op = 1
        frmMain.Command8_Click
        If frmMain.txtOperations.Text = "1" Then frmMain.txtOperations.Text = ""
    Loop
    
    frmMain.Command8_Click
    If frmMain.txtOperations.Text = "1" Then frmMain.txtOperations.Text = ""
    
    Header = Header & "\tab Rate\tab Total"
    Header = Right(Header, Len(Header) - 5)
    Temp = Replace(Temp, "Headers", Header)
    Temp = Replace(Temp, "OutBal", Format(CDbl(frmMain.txtFields(4).Text), "$#0.00"))
    Temp = Replace(Temp, "CurBal", Format(Total, "$#0.00"))
    Temp = Replace(Temp, "TotBal", Format((CDbl(frmMain.txtFields(4).Text) + Total), "$#0.00"))
    
    Text1.TextRTF = Temp
    
    Text1.SelStart = 1
    Text1.OLEObjects.Add , , App.Path & "\Color Logo.bmp"
    Text1.SelStart = 1
    
    frmMain.txtFields(4).Text = Format(CDbl(frmMain.txtFields(4).Text), "$#0.00")
    frmMain.txtFields(1).Text = Format(Total, "$#0.00")
    frmMain.txtFields(6).Text = Format((CDbl(frmMain.txtFields(4).Text) + Total), "$#0.00")
    
    Me.Text1.SaveFile App.Path & "\Backups\" & frmMain.txtFields(0).Text & "= " & Format(Date, "mm-yyyy") & ".rtf"
End Sub

