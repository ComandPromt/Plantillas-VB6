VERSION 5.00
Object = "{9B1E48ED-8018-11D3-B75D-006097A1EBF0}#1.0#0"; "DNS.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmActionWho 
   Caption         =   "ActionWho 2000"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   Icon            =   "frmActionWho.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog mcdCommon 
      Left            =   60
      Top             =   945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.*"
      DialogTitle     =   "Öpnna textfil"
      Filter          =   "*.*"
      InitDir         =   "C:\"
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "List"
      Height          =   405
      Left            =   3135
      TabIndex        =   6
      Top             =   1035
      Width           =   1215
   End
   Begin DNSControl.DNS dns1 
      Left            =   30
      Top             =   1005
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton cmdCommand2 
      Caption         =   "Name to IP"
      Height          =   405
      Left            =   1845
      TabIndex        =   3
      Top             =   1035
      Width           =   1215
   End
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "IP to Name"
      Height          =   405
      Left            =   555
      TabIndex        =   2
      Top             =   1035
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   345
      Left            =   525
      TabIndex        =   1
      Top             =   555
      Width           =   4485
   End
   Begin VB.TextBox txtIP 
      Height          =   330
      Left            =   525
      TabIndex        =   0
      Top             =   105
      Width           =   4485
   End
   Begin VB.Label lblCounter 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   240
      Left            =   4410
      TabIndex        =   7
      Top             =   1125
      Width           =   570
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   285
      Left            =   30
      TabIndex        =   5
      Top             =   630
      Width           =   660
   End
   Begin VB.Label lblIP 
      Caption         =   "IP:"
      Height          =   285
      Left            =   30
      TabIndex        =   4
      Top             =   180
      Width           =   660
   End
End
Attribute VB_Name = "frmActionWho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCommand1_Click()
    txtName.Text = dns1.AddressToName(txtIP.Text)
End Sub
Private Sub cmdCommand2_Click()
    txtIP.Text = dns1.NameToAddress(txtName.Text)
End Sub
Private Sub cmdList_Click()
    Dim cFileIn As String
    Dim cFileOut As String
    Dim cIPnumber As String
    Dim cDNSname As String
    Dim lCounter As Long
    mcdCommon.FileName = ""
    mcdCommon.ShowOpen
    frmActionWho.Refresh
    If Len(Trim(mcdCommon.FileName)) > 0 Then
        cFileIn = mcdCommon.FileName
        cFileOut = Left(mcdCommon.FileName, Len(mcdCommon.FileName) - 4) + " searched.txt"
        Open cFileIn For Input As #1
        Open cFileOut For Output As #2
        While Not EOF(1)
            Input #1, cIPnumber
            txtIP.Text = Trim(cIPnumber)
            cDNSname = Trim(dns1.AddressToName(Trim(cIPnumber)))
            txtName.Text = Trim(cDNSname)
            Write #2, cIPnumber, cDNSname
            lCounter = lCounter + 1
            lblCounter.Caption = lCounter
            lblCounter.Refresh
            DoEvents
        Wend
        Close
        MsgBox "Ready, " & lCounter & " lines searched.", vbInformation, "ActionWho 2000"
        lblCounter = 0
        txtIP.Text = ""
        txtName.Text = ""
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Close
End Sub
