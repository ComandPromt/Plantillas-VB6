VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Binary To Decimal"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBintoDec 
      Caption         =   ">"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmdDectoBin 
      Caption         =   ">"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox txtBinary 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtDecimal 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblBinary 
      AutoSize        =   -1  'True
      Caption         =   "Binary Number  :"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label lblDecimal 
      AutoSize        =   -1  'True
      Caption         =   "Decimal Number :"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Function DecToBin(ByVal x As Long) As String
Dim Y As Long
Dim Num As String

While (x \ 2) > 0
    Y = x \ 2
    If x > 1 Then
    Num = Val(x - (Y * 2)) & Num
    End If
    x = Y
Wend
Num = x & Num
DecToBin = Num
End Function
Private Function Pow(Number As Long, Power As Integer) As Long
Dim x As Integer
If Power > 1 Then
For x = 2 To Power
Number = Number * 2
Next
Pow = Number
Else
If Power = 0 Then Pow = 1
If Power = 1 Then Pow = Number
End If
End Function
Private Function BinToDec(ByVal Num As Long) As Long
Dim sngNumber As Long
Dim x As Integer
Dim Tmp As Long
Dim Output As Long
sngNumber = Num

For x = 0 To Len(CStr(sngNumber)) - 1
Tmp = CLng(Right(CStr(sngNumber), 1))
If Tmp = 1 Then
Tmp = Tmp * Pow(2, x)
End If
Output = Output + Tmp
Tmp = 1

If Len(CStr(sngNumber)) > 1 Then
sngNumber = CLng(Left$(CStr(sngNumber), Len(CStr(sngNumber)) - 1))
Else
sngNumber = 0
End If
Next
BinToDec = Output
End Function


Private Sub cmdBintoDec_Click()
txtDecimal.Text = BinToDec(txtBinary.Text)
txtBinary.Text = ""
End Sub

Private Sub cmdDectoBin_Click()
txtBinary.Text = DecToBin(txtDecimal.Text)
txtDecimal.Text = ""
End Sub


