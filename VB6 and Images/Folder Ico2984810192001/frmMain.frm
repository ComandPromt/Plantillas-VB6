VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon changer"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   4665
      TabIndex        =   4
      Top             =   3465
      Width           =   1500
   End
   Begin VB.CommandButton cmdChange 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      Picture         =   "frmMain.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Icons"
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   6855
      Begin VB.Frame Frame3 
         Caption         =   "New"
         Height          =   1815
         Left            =   4440
         TabIndex        =   8
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton cmdSelectFile 
            Caption         =   "Select file..."
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   1200
            Width           =   1215
         End
         Begin VB.HScrollBar hscIcons 
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1200
            Width           =   1215
         End
         Begin VB.PictureBox picNew 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   240
            ScaleHeight     =   61
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   77
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Current:"
         Height          =   1815
         Left            =   975
         TabIndex        =   7
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton cmdDefault 
            Caption         =   "Set to &default"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   1
            Top             =   1200
            Width           =   1215
         End
         Begin VB.PictureBox picCurrent 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   240
            ScaleHeight     =   61
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   77
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.CommandButton cmdSelectFolder 
      Caption         =   "Select  folder"
      Height          =   495
      Left            =   6390
      TabIndex        =   0
      Top             =   120
      Width           =   600
   End
   Begin MSComDlg.CommonDialog dlgFolder 
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblFolder 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ICON CHANGER
'You can distribute and modify this program as you wish, as long as original credit stays here
'Programmed by Janeks Bergs 2001
'mailto: janexx@kuldiga.lv

Option Explicit

Public FolderSelected As String

'structure to hold all text lines of INI file and starting and ending positions of "[.ShellClassInfo]" tag
Private Type INIfileType
   Lines() As String
   TagStartLine As Integer
   TagEndLine As Integer
   IconFileLine As Integer
   IconIndexLine As Integer
End Type

Dim DesktopINI As INIfileType    'will be used for reading INI file and editing it afterwards
Dim EmptyINI As INIfileType      'to quickly clear DesktopINI file

Dim IconFile As String      'self explanatory
Dim IconIndex As Long

'Windows API functions needed for extracting and drawing icons
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Sub cmdChange_Click()
   Dim IniFile As String
   Dim LinesArrayUbound As Integer
   Dim t As Integer
   Dim f As Integer
   Dim NoDesktopINI As Boolean
   
   f = FreeFile
   IniFile = FolderSelected & "\desktop.ini"
   
   If FileExists(IniFile) Then NoDesktopINI = False Else NoDesktopINI = True
   
   Open IniFile For Output As #f
   
   If NoDesktopINI Then
      'if there is no INI file - create new file with tag and values
      ReDim DesktopINI.Lines(3)
      DesktopINI.Lines(1) = "[.ShellClassInfo]"
      DesktopINI.Lines(2) = "IconFile=" & Chr$(34) & picNew.Tag & Chr$(34)
      DesktopINI.Lines(3) = "IconIndex=" & hscIcons.Value
   ElseIf Not NoDesktopINI And DesktopINI.TagStartLine = 0 Then
      'INI file exists, but there is no [.ShellClassInfo] tag - append tag and values
      LinesArrayUbound = UBound(DesktopINI.Lines)
      ReDim Preserve DesktopINI.Lines(LinesArrayUbound + 3)
      DesktopINI.Lines(LinesArrayUbound + 1) = "[.ShellClassInfo]"
      DesktopINI.Lines(LinesArrayUbound + 2) = "IconFile=" & Chr$(34) & picNew.Tag & Chr$(34)
      DesktopINI.Lines(LinesArrayUbound + 3) = "IconIndex=" & hscIcons.Value
   ElseIf Not NoDesktopINI And DesktopINI.TagStartLine <> 0 And DesktopINI.IconFileLine = 0 Then
      'INI file exists and tag also exists, but there is no info about icon file and index - insert values in tag
      LinesArrayUbound = UBound(DesktopINI.Lines)
      ReDim Preserve DesktopINI.Lines(LinesArrayUbound + 2)
      For t = LinesArrayUbound To DesktopINI.TagStartLine + 1 Step -1   'move all text lines in array 2 indexes higher - free space for inserting values
         DesktopINI.Lines(t + 2) = DesktopINI.Lines(t)
      Next t
      DesktopINI.Lines(DesktopINI.TagStartLine + 1) = "IconFile=" & Chr$(34) & picNew.Tag & Chr$(34)
      DesktopINI.Lines(DesktopINI.TagStartLine + 2) = "IconIndex=" & hscIcons.Value
   ElseIf Not NoDesktopINI And DesktopINI.TagStartLine <> 0 And DesktopINI.IconFileLine <> 0 Then
      'INI file exists, tag exists and icon info also exists - edit existing values
      DesktopINI.Lines(DesktopINI.IconFileLine) = "IconFile=" & Chr$(34) & picNew.Tag & Chr$(34)
      DesktopINI.Lines(DesktopINI.IconIndexLine) = "IconIndex=" & hscIcons.Value
   End If
   
   LinesArrayUbound = UBound(DesktopINI.Lines)
   For t = 1 To LinesArrayUbound
      Print #1, DesktopINI.Lines(t)
   Next t
   
   Close #f
   
   SetAttr FolderSelected, vbSystem
   SetAttr IniFile, vbHidden
   
   LoadFolderData
   
End Sub

Private Sub cmdDefault_Click()
   Dim NoDesktopINI As Boolean
   Dim IniFile As String
   Dim t As Integer
   Dim f As Integer
   Dim TrueBound As Integer
   
   Set picCurrent.Picture = LoadPicture("")
   picCurrent.Refresh
   picCurrent.PaintPicture LoadResPicture("DEFAULT", 0), 23, 9
   
   IniFile = FolderSelected & "\desktop.ini"
   
   If FileExists(IniFile) Then NoDesktopINI = False Else NoDesktopINI = True
   
   If NoDesktopINI Then
      SetAttr FolderSelected, vbNormal
   ElseIf Not NoDesktopINI And DesktopINI.TagStartLine = 0 Then
      SetAttr FolderSelected, vbNormal
   ElseIf Not NoDesktopINI And DesktopINI.TagStartLine <> 0 And DesktopINI.IconFileLine = 0 Then
      'if there are only 1 line, it's tag line, so delete file
      If UBound(DesktopINI.Lines) = 1 Then
         SetAttr IniFile, vbNormal
         Kill (IniFile)
      End If
      SetAttr FolderSelected, vbNormal
   ElseIf Not NoDesktopINI And DesktopINI.TagStartLine <> 0 And DesktopINI.IconFileLine <> 0 Then
      
      If UBound(DesktopINI.Lines) = 3 Then
         'if there are 3 lines, they are tag and icon info lines, so we can delete whole file
         SetAttr IniFile, vbNormal
         Kill (IniFile)
      Else
         'remove icon information
         TrueBound = UBound(DesktopINI.Lines)
         ReDim Preserve DesktopINI.Lines(TrueBound + 2) 'avoid errors in loop afterwards
         For t = DesktopINI.IconFileLine To TrueBound
            DesktopINI.Lines(t) = DesktopINI.Lines(t + 2)
         Next t
         ReDim Preserve DesktopINI.Lines(UBound(DesktopINI.Lines) - 4)  'remove empty strings
         
         f = FreeFile
         Open IniFile For Output As #f
         
         For t = 1 To UBound(DesktopINI.Lines)
            Print #f, DesktopINI.Lines(t)
         Next t
         
         Close #f
      End If
      
      SetAttr FolderSelected, vbNormal
      
   End If
   
   LoadFolderData
   cmdDefault.Enabled = False
   
End Sub

Private Sub cmdQuit_Click()
   Unload Me
End Sub

Private Sub cmdSelectFolder_Click()
   Dim Backup As String
   
   Backup = FolderSelected
   
   FolderSelected = ""
   
   frmFolder.Move Me.Left + Me.Width, Me.Top
   frmFolder.Dir1.Path = lblFolder.Caption
   frmFolder.Show vbModal
   
   If FolderSelected = "" Then
      FolderSelected = Backup
      Exit Sub
   Else
      LoadFolderData
   End If
   
End Sub

Private Sub lblFolder_Change()
   If TextWidth(lblFolder.Caption) > lblFolder.Width Then
      lblFolder.Caption = Left$(lblFolder.Caption, Len(lblFolder.Caption) - 4) & "..."
   End If
End Sub

Private Sub LoadFolderData()
   Dim IniPath As String
   Dim f As Integer
   Dim c As Integer
   Dim IsInTag As Boolean
   
   IniPath = FolderSelected & "\desktop.ini"
   DesktopINI = EmptyINI
   IconFile = "": IconIndex = 0
   
   f = FreeFile
   If FileExists(IniPath) Then
      Open IniPath For Input As #f
      Do
         c = c + 1
         ReDim Preserve DesktopINI.Lines(c)
         Line Input #f, DesktopINI.Lines(c)
         If InStr(1, DesktopINI.Lines(c), "[.ShellClassInfo]") <> 0 Then
            IsInTag = True
            DesktopINI.TagStartLine = c
         End If
         If IsInTag Then
            If InStr(1, DesktopINI.Lines(c), "[") <> 0 And InStr(1, DesktopINI.Lines(c), "]") <> 0 And _
            c <> DesktopINI.TagStartLine Then   'check for start of new tag and exclude detection of starting line
               DesktopINI.TagEndLine = c - 1 'end line is previous
               IsInTag = False
            End If
            If InStr(1, DesktopINI.Lines(c), "IconFile=") <> 0 Then
               IconFile = Mid$(DesktopINI.Lines(c), InStr(1, DesktopINI.Lines(c), "IconFile=") + 9)
               IconFile = RemoveChar(IconFile, Chr$(34))
               DesktopINI.IconFileLine = c
            ElseIf InStr(1, DesktopINI.Lines(c), "IconIndex=") <> 0 Then
               IconIndex = Mid$(DesktopINI.Lines(c), InStr(1, DesktopINI.Lines(c), "IconIndex=") + 10)
               DesktopINI.IconIndexLine = c
            End If
         End If
         If EOF(f) Then
            If IsInTag Then
               DesktopINI.TagEndLine = c
               IsInTag = False
            End If
            Exit Do
         End If
      Loop
      
      If IconFile <> "" Then
         If PutIcon(picCurrent, IconFile, IconIndex) = False Then
            MsgBox "Error while extracting icon #" & IconIndex & " from file:" & vbCrLf & IconFile
            cmdDefault.Enabled = False
         Else
            cmdDefault.Enabled = True
            If picNew.Tag <> "" Then cmdChange.Enabled = True
         End If
      Else
         Set picCurrent.Picture = LoadPicture("")
         picCurrent.PaintPicture LoadResPicture("DEFAULT", 0), 23, 9
         If picNew.Tag <> "" Then cmdChange.Enabled = True
      End If
      
   Else
      Set picCurrent.Picture = LoadPicture("")
      picCurrent.PaintPicture LoadResPicture("DEFAULT", 0), 23, 9
      If picNew.Tag <> "" Then cmdChange.Enabled = True
   End If
      
   Close #f
   
End Sub

Private Function PutIcon(pic As PictureBox, IconFile As String, IconIndex As Long) As Boolean
   'puts icon in PictureBox
   
   Dim hIcon As Long  'operating system's handle of icon object
   
   If UCase$(Right$(IconFile, 4) = ".ICO") Then
      'it's icon file
      Set pic.Picture = LoadPicture("")  'clear picture box
      pic.PaintPicture LoadPicture(IconFile), 23, 14  'draw icon centered
      pic.Refresh
      
      PutIcon = True
   Else
      ExtractIconEx IconFile, IconIndex, hIcon, ByVal 0&, 1
      
      If hIcon = 0 Then
         'can't extract icon
         PutIcon = False
         Exit Function
      Else
         'if everything was ok, draw icon in the middle of picturebox
         Set pic.Picture = LoadPicture("")
         DrawIcon pic.hdc, 23, 14, hIcon
         '...and destroy its handle
         DestroyIcon hIcon
         
         pic.Refresh
         PutIcon = True
      End If
   End If
      
End Function
Private Function FileExists(Path As String) As Boolean
   Dim f As Integer
   On Error Resume Next
   
   f = FreeFile
   
   Open Path For Input As #f
   If Err.Number = 53 Then
      FileExists = False
   Else
      If FileLen(Path) = 0 Then
         Close #f
         SetAttr Path, vbNormal
         Kill (Path)
         FileExists = False
         Exit Function
      End If
      
      FileExists = True
   End If
   Close #f
   
End Function

Private Function RemoveChar(ByVal Source As String, Char As String)
   'this function removes given character from string
   Dim t As Integer
   
   For t = 1 To Len(Source)
      If Mid$(Source, t, 1) = Char Then
         Source = Left$(Source, t - 1) & Mid$(Source, t + 1)
         t = t - 1
      End If
   Next t
   
   RemoveChar = Source
   
End Function

Private Sub cmdSelectFile_Click()
   On Error GoTo Cancel
   Dim success As Boolean
      
   With dlgFolder
      .CancelError = True
      .Filter = "All icon containing files|*.ico;*.exe;*.dll;*.icl;*.ocx;|Icons (*.ico)|*.ico|Icon Libraries (*.icl)|*.icl|Executables (*.exe)|*.exe|Dynamic Link Libraries (*.dll)|*.dll|ActiveX controls (*.ocx)|*.ocx"
      .FilterIndex = 1
      .ShowOpen
   End With
   
   success = PutIcon(picNew, dlgFolder.FileName, 0)
   If success Then
      picNew.Tag = dlgFolder.FileName
      
      cmdSelectFile.Top = 1440
      hscIcons.Min = 0
      hscIcons.Max = ExtractIconEx(dlgFolder.FileName, -1, ByVal 0&, ByVal 0&, ByVal 0&) - 1
      hscIcons.Visible = True: hscIcons.Value = 0
      If FolderSelected <> "" Then cmdChange.Enabled = True
   Else
      MsgBox "Selected file contains no icons!", vbCritical, "Error"
      cmdChange.Enabled = False
      hscIcons.Min = 0: hscIcons.Max = 0
      Set picNew.Picture = LoadPicture("")
      picNew.Tag = ""
   End If
   
Cancel:
   
End Sub

Private Sub hscIcons_Change()
   Dim success As Boolean
   success = PutIcon(picNew, picNew.Tag, hscIcons.Value)
   
   If Not success Then
      MsgBox "Error while extracting icon", vbCritical, "Error"
   End If
   
End Sub
