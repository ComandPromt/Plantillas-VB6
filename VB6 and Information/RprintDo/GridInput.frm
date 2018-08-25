VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Formiga 
   BackColor       =   &H00FFFFFF&
   Caption         =   "GridEdit "
   ClientHeight    =   8460
   ClientLeft      =   255
   ClientTop       =   1005
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   11250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   -240
      Width           =   9855
      Begin VB.TextBox Invoice 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6090
         TabIndex        =   11
         Text            =   "Invoice"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Country 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6120
         TabIndex        =   9
         Text            =   "Country"
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox CustAddress 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6090
         TabIndex        =   8
         Text            =   "CustAddress"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox Date 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6090
         TabIndex        =   7
         Text            =   "Date"
         Top             =   300
         Width           =   1935
      End
      Begin VB.TextBox City 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Text            =   "City"
         Top             =   1680
         Width           =   2580
      End
      Begin VB.TextBox CustomerName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   3480
         TabIndex        =   4
         Text            =   "CustomerName"
         Top             =   1080
         Width           =   2565
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   3840
         X2              =   7680
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MsFlexGrid Edit Cell Example"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   15
         TabIndex        =   12
         Tag             =   "noprint"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   1335
         Left            =   -15
         Picture         =   "GridInput.frx":0000
         Stretch         =   -1  'True
         Top             =   135
         Width           =   3000
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Order"
         Height          =   255
         Left            =   5340
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill To"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   225
      TabIndex        =   0
      Top             =   2175
      Width           =   9975
      Begin VB.PictureBox inputpic 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1065
         ScaleHeight     =   270
         ScaleWidth      =   795
         TabIndex        =   1
         Tag             =   "noprint"
         Top             =   735
         Visible         =   0   'False
         Width           =   825
         Begin VB.TextBox TMText 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   15
            TabIndex        =   2
            ToolTipText     =   "Enter data"
            Top             =   15
            Width           =   840
         End
      End
      Begin MSFlexGridLib.MSFlexGrid gridinv 
         Height          =   4965
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   8758
         _Version        =   393216
         Rows            =   25
         Cols            =   5
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Total 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7290
         TabIndex        =   14
         Top             =   5490
         Width           =   615
      End
      Begin VB.Label TotalInv 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   8025
         TabIndex        =   13
         Top             =   5505
         Width           =   1110
      End
   End
End
Attribute VB_Name = "Formiga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastLine As Integer
Dim Edit As Boolean
Dim OldCol As Integer
Dim OldRow As Integer

Private Sub Form_Activate()
gridinv.Col = 0
End Sub

Private Sub Form_Load()
LastLine = 25
Setgrid
gridinv.Col = 1
End Sub

Private Sub gridinv_EnterCell()
Dim ResNo As Integer
On Error Resume Next
If Init Then Exit Sub
ResNo = 116
'Total Column Not allowed for write'
If gridinv.Col = 4 Then
gridinv.Col = OldCol
Exit Sub
End If
'Set position and size of the EditText unit
inputpic.Move gridinv.CellLeft - 20
inputpic.Top = gridinv.CellTop + gridinv.Top - 20
inputpic.Visible = True

inputpic.Width = gridinv.CellWidth + 10
inputpic.Height = gridinv.CellHeight + 20

TMText.Width = inputpic.Width
'Keep old position
OldCol = gridinv.Col
OldRow = gridinv.Row

'Select the new cell Text
TMText.SetFocus
TMText.Text = gridinv.Text
TMText.SelStart = 0
TMText.SelLength = 100
End Sub

Private Sub GridInv_KeyDown(KeyCode As Integer, Shift As Integer)
Dim reik
reik = ""
Select Case KeyCode
Case KEY_HOME
gridinv.Row = 1
KeyCode = 0
If Not gridinv.RowIsVisible(1) Then gridinv.TopRow = 1
Case KEY_END
KeyCode = 0
gridinv.Row = LastLine
If Not gridinv.RowIsVisible(LastLine) Then gridinv.TopRow = LastLine - 3
Case KEY_HOME
gridinv.Row = 1:
KeyCode = 0
If Not gridinv.RowIsVisible(1) Then gridinv.TopRow = 1
End Select
If KeyCode = 46 Then
If Not Edit Then
gridinv.Text = ""
Else

Exit Sub
End If
End If
Newline = 0
KeyCode = 0

End Sub

Private Sub gridinv_LeaveCell()
Dim The_Col As Integer, the_Row As Integer
Dim Tot As Double
If gridinv.Col = 4 Then Exit Sub
The_Col = gridinv.Col
the_Row = gridinv.Row
gridinv.Text = TMText.Text

If Edit And (The_Col = 2 Or The_Col = 3) And (Not IsNumeric(gridinv.Text)) Then gridinv.Text = "0"
If Edit And (The_Col = 2 Or The_Col = 3) And (IsNumeric(gridinv.TextMatrix(the_Row, 2)) _
And IsNumeric(gridinv.TextMatrix(the_Row, 3))) Then
Tot = CDbl(gridinv.TextMatrix(the_Row, 2)) * CDbl(gridinv.TextMatrix(the_Row, 3))
gridinv.TextMatrix(the_Row, 4) = Str$(Tot)
TotalSum
End If
Edit = False
 
End Sub

Private Sub gridinv_Scroll()

If Not gridinv.RowIsVisible(gridinv.Row) Or Not gridinv.ColIsVisible(gridinv.Col) Then
inputpic.Visible = False
Else
inputpic.Move gridinv.CellLeft
inputpic.Top = gridinv.CellTop + gridinv.Top
inputpic.Visible = True
End If

End Sub

Private Sub TMtext_KeyDown(KeyCode As Integer, Shift As Integer)
Edit = True
Select Case KeyCode
Case 27
TMText = gridinv.Text
Case 40
If gridinv.Row < LastLine - 1 Then gridinv.Row = gridinv.Row + 1
KeyCode = 0
Case 39 'Right Arrow action
KeyCode = 0
If gridinv.Col < 3 Then
gridinv.Col = gridinv.Col + 1
ElseIf gridinv.Row < gridinv.Rows - 1 Then
gridinv.Row = gridinv.Row + 1
gridinv.Col = 0

End If
Case 37 'Left Arrow Action
KeyCode = 0 'Clean the buffer
If gridinv.Col > 0 Then
gridinv.Col = gridinv.Col - 1
ElseIf gridinv.Row > 1 Then
gridinv.Row = gridinv.Row + -1
gridinv.Col = 3
End If
Case 38
If gridinv.Row > 1 Then gridinv.Row = gridinv.Row - 1
KeyCode = 0
Case KEY_END
KeyCode = 0
gridinv.Row = LastLine
If Not gridinv.RowIsVisible(LastLine) Then gridinv.TopRow = LastLine - 3
Case KEY_HOME
gridinv.Row = 1:
KeyCode = 0
If Not gridinv.RowIsVisible(1) Then gridinv.TopRow = 1
End Select
If KeyCode = 46 Then
If Not Edit Then
gridinv.Text = ""

Exit Sub
End If
End If
TMText.SetFocus
End Sub

Private Sub TMtext_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
If gridinv.Col < 3 Then
gridinv.Col = gridinv.Col + 1
Else
If gridinv.Row < LastLine Then gridinv.Col = 0: gridinv.Row = gridinv.Row + 1
End If
Else

Edit = True
End If
End Sub

Private Sub Setgrid()
Dim Dollars As Currency
Dim Qntys As Double
gridinv.Move 0, 300
inputpic.Visible = True
gridinv.ColWidth(1) = 3600
gridinv.ColWidth(2) = 1500
gridinv.ColWidth(3) = 1500
gridinv.ColWidth(4) = 1500
gridinv.TextMatrix(0, 0) = "Item Code"
gridinv.TextMatrix(0, 1) = "Item Name"
gridinv.TextMatrix(0, 2) = "Quantity"
gridinv.TextMatrix(0, 3) = "Price"
gridinv.TextMatrix(0, 4) = "SubTotal"
gridinv.TextMatrix(1, 0) = "0007845"
gridinv.TextMatrix(1, 1) = "Giraffe"
gridinv.ColAlignment(3) = 8
gridinv.ColAlignment(4) = 8
gridinv.TextMatrix(1, 2) = "35"
Dollars = 118.95
gridinv.TextMatrix(1, 3) = Str$(Dollars)
Dollars = 4163.25
gridinv.TextMatrix(1, 4) = Str$(Dollars)
gridinv.TextMatrix(2, 0) = "0006984"
gridinv.TextMatrix(2, 1) = "Lion"
gridinv.TextMatrix(2, 2) = "7"
Dollars = 99.99
gridinv.TextMatrix(2, 3) = Str$(Dollars)
Dollars = 349.93
gridinv.TextMatrix(2, 4) = Str$(Dollars)
gridinv.TextMatrix(3, 0) = "0000183"
gridinv.TextMatrix(3, 1) = "Snake"
gridinv.TextMatrix(3, 2) = "1"
Dollars = 96.99
gridinv.TextMatrix(3, 3) = Str$(Dollars)
gridinv.TextMatrix(3, 4) = Str$(Dollars)
gridinv.TextMatrix(4, 0) = "0003155"
gridinv.TextMatrix(4, 1) = "Hippopotami"
gridinv.TextMatrix(4, 2) = "10"
Dollars = 214.99
gridinv.TextMatrix(4, 3) = Str$(Dollars)
gridinv.TextMatrix(4, 4) = Str$(Dollars * 10)
gridinv.TextMatrix(5, 0) = "0007825"
gridinv.TextMatrix(5, 1) = "Crocodile"
gridinv.TextMatrix(5, 2) = "6"
Dollars = 235.99
gridinv.TextMatrix(5, 3) = Str$(Dollars)
gridinv.TextMatrix(5, 4) = Str$(Dollars * 6)
gridinv.Col = 0
LastLine = 6

gridinv_EnterCell
Date = Format(Now, "dd/mm/yy")
Invoice = "1000345"
CustomerName = "Santa Claus"
CustAddress = "Sky st 1"
City = "Los Angeles"
Country = "Holyland"
TotalSum
TMText.SelStart = 0
TMText.SelLength = 100
End Sub

Private Sub TotalSum()
Dim i As Integer
Dim Sum As Double
Dim IsEmpty As Boolean
i = 1
Do While Not IsEmpty
If gridinv.TextMatrix(i, 4) = "" Then
IsEmpty = True
Else
Sum = Sum + CDbl(gridinv.TextMatrix(i, 4))
i = i + 1
End If
Loop
TotalInv = Str$(Sum)
End Sub

Private Sub TMText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
TMText.SelLength = 0
End Sub
