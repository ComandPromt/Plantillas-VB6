VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   Caption         =   "DragonBall Browser - Favorites"
   ClientHeight    =   4230
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1005
      Left            =   210
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   1773
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form2.frx":0442
   End
   Begin VB.ListBox Favorites 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      ItemData        =   "Form2.frx":050A
      Left            =   45
      List            =   "Form2.frx":050C
      TabIndex        =   0
      Top             =   60
      Width           =   3885
   End
   Begin VB.Menu Comm 
      Caption         =   "C&ommands"
      Begin VB.Menu LList 
         Caption         =   "&Load List"
      End
      Begin VB.Menu SList 
         Caption         =   "&Save List"
      End
      Begin VB.Menu pokpfs 
         Caption         =   "-"
      End
      Begin VB.Menu AList 
         Caption         =   "&Add To List"
         Shortcut        =   ^A
      End
      Begin VB.Menu RList 
         Caption         =   "&Remove From List"
         Shortcut        =   ^R
      End
      Begin VB.Menu poksdfpok 
         Caption         =   "-"
      End
      Begin VB.Menu Cl 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Sub AList_Click()
Dim ToAddName As String, ToAddURL As String
ToAddName$ = InputBox("Enter the name of the site <not the URL>", "DragonBall Browser v1.0.0")
If ToAddName$ = "" Then
    MsgBox "To add a site to the list, you must enter a valid name.", vbInformation, "DragonBall Browser v1.0.0"
    Exit Sub
End If
ToAddURL$ = InputBox("Enter the URL of the site <not the name>", "DragonBall Browser v1.0.0")
If ToAddURL$ = "" Then
    MsgBox "To add a site to the list, you must enter a valid URL.", vbInformation, "DragonBall Browser v1.0.0"
    Exit Sub
End If
Favorites.AddItem ToAddName$ & " - " & ToAddURL$
End Sub

Private Sub Cl_Click()
Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
Favorites.Width = Form2.Width - 210: Favorites.Height = Form2.Height - (4905 - 4110)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Response
Response = MsgBox("Are you sure you wish to exit? Unsaved data will be lost...", vbInformation Or vbYesNo, "DragonBall Browser v1.0.0")
If Response = vbNo Then Cancel = 1
If Response = vbYes Then Unload Me
End Sub

Private Sub LList_Click()
On Error GoTo ErrHandler
Dim NumOfFavs As Integer, OrLen As Integer, Temporary As Integer, Splitter() As String
RichTextBox1.LoadFile App.Path & "\Favorites.dbz"
OrLen = Len(RichTextBox1.Text): NumOfFavs = Mid(RichTextBox1.Text, 1, 1)
Temporary = InStr(RichTextBox1.Text, vbNewLine): Temporary = Tenporary + 4
RichTextBox1.Text = Mid(RichTextBox1.Text, Temporary, Len(RichTextBox1.Text))
Splitter = Split(RichTextBox1.Text, "|"): RichTextBox1.Text = "": Favorites.Clear
For i = 0 To NumOfFavs - 1
    Favorites.AddItem Splitter(i)
Next i
Exit Sub
ErrHandler: MsgBox Err.Description, vbInformation, "DragonBall Browser v1.0.0": RichTextBox1.Text = ""
End Sub

Private Sub RList_Click()
Dim NumToRemove
NumToRemove = InputBox("Enter the number of the item you wish to remove", "DragonBall Browser v1.0.0")
If NumToRemove = 0 Then Exit Sub: If NumToRemove = "" Then Exit Sub
If NumToRemove > Favorites.ListCount Then
    MsgBox "Please enter a valid number to remove.", vbInformation, "DragonBall Browser v1.0.0"
    Exit Sub
End If
Favorites.RemoveItem (NumToRemove - 1)
End Sub

Private Sub SList_Click()
On Error GoTo ErrHandler
RichTextBox1.Text = Favorites.ListCount & vbNewLine
For i = 0 To Favorites.ListCount - 1
    RichTextBox1.Text = RichTextBox1.Text & Favorites.List(i) & "|"
Next i
RichTextBox1.SaveFile App.Path & "\Favorites.dbz": Exit Sub
ErrHandler: MsgBox Err.Description, vbinforamtion, "DragonBall Browser v1.0.0": RichTextBox1.Text = ""
End Sub

