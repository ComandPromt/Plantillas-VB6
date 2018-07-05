VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmText 
   Caption         =   "Text Picture Box"
   ClientHeight    =   4470
   ClientLeft      =   75
   ClientTop       =   1770
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   6375
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
   Begin VB.Timer Cursor 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   120
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileSub 
         Caption         =   "Clear Text"
         Index           =   0
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "Change Font"
         Index           =   1
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "Print"
         Index           =   2
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "E&xit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CursorOn As Boolean, DrawObj As Object

Private Sub Form_Load()
    Set DrawObj = Picture1
    Cursor.Enabled = True
End Sub

Private Sub mnuFileSub_Click(Index As Integer)
    Select Case Index
        Case 0 ' Clear Text
            DrawObj.Cls
        Case 1 ' Change Font Properties
            Cursor.Enabled = False
            If CursorOn Then SetCursor
            With CommonDialog1
                .Flags = cdlCFBoth Or cdlCFEffects
                .FontName = DrawObj.FontName
                .FontItalic = DrawObj.FontItalic
                .FontSize = DrawObj.FontSize
                .FontStrikethru = DrawObj.FontStrikethru
                .FontBold = DrawObj.FontBold
                .FontUnderline = DrawObj.FontUnderline
                .ShowFont
                If Not .CancelError Then
                    DrawObj.FontName = .FontName
                    DrawObj.FontItalic = .FontItalic
                    DrawObj.FontSize = .FontSize
                    DrawObj.FontStrikethru = .FontStrikethru
                    DrawObj.FontBold = .FontBold
                    DrawObj.FontUnderline = .FontUnderline
                    DrawObj.ForeColor = .Color
                End If
            End With
            Cursor.Enabled = True
        Case 2 'Print
            Cursor.Enabled = False
            If CursorOn Then SetCursor
            Printer.PaintPicture DrawObj.Image, 0, 0
            Printer.EndDoc
            Cursor.Enabled = True
        Case 3
            End
    End Select
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
    Cursor.Enabled = False
    If CursorOn Then SetCursor
    If KeyAscii = 13 Then DrawObj.Print "" Else DrawObj.Print Chr(KeyAscii);
    Cursor.Enabled = True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Cursor.Enabled = False
    If CursorOn Then SetCursor
    DrawObj.CurrentX = X
    DrawObj.CurrentY = Y
    Cursor.Enabled = True
    If Button = 2 Then
    PopupMenu mnuFile
    End If
    
End Sub
Private Sub SetCursor()
    DrawObj.DrawMode = 6
    SaveCurrentY = DrawObj.CurrentY
    SaveCurrentX = DrawObj.CurrentX
    CursorHeight = DrawObj.TextHeight("I")
    DrawObj.Line (SaveCurrentX, SaveCurrentY)-(SaveCurrentX, SaveCurrentY + CursorHeight)
    DrawObj.CurrentY = SaveCurrentY
    DrawObj.CurrentX = SaveCurrentX
    DrawObj.DrawMode = 13
    CursorOn = Not CursorOn
End Sub
Private Sub Cursor_Timer()
    SetCursor
End Sub

