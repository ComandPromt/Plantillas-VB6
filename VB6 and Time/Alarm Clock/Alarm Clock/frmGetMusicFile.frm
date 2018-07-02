VERSION 5.00
Begin VB.Form frmGetMusicFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Music File..."
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ComboBox cboTypeChange 
      Height          =   315
      ItemData        =   "frmGetMusicFile.frx":0000
      Left            =   2400
      List            =   "frmGetMusicFile.frx":000A
      TabIndex        =   3
      Text            =   "*.mp3"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.FileListBox filSelectFile 
      Height          =   1455
      Left            =   2400
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.DirListBox dirSelectDir 
      Height          =   1440
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.DriveListBox drvSelectDrive 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmGetMusicFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboTypeChange_Click()

' Input: None
' Process: When the text in the combo-box changes, the file pattern
'          changes with it, and the file list changes to reflect the
'          new pattern
' Output: None

    On Error GoTo errorhandler ' Catching any errors in the following code
    filSelectFile.Pattern = cboTypeChange.Text ' Changing the pattern
    filSelectFile.Refresh ' Refreshing the file list
    Exit Sub
errorhandler: ' An error was found
    MsgBox "Error!  File type not allowed!", 16, "Error!"
End Sub

Private Sub cmdOK_Click()

' Input: None
' Process: Constructing the complete file path and storing it in a
'          global variable
' Output: String variable "filename1"

    If Len(dirSelectDir.Path) > 3 Then ' Checking if the file is in a root
        filename1 = filSelectFile.Path & "\" & filSelectFile.filename ' It's not
    Else
        filename1 = filSelectFile.Path & filSelectFile.filename ' It is
    End If
    Unload Me ' Unloading the form
End Sub

Private Sub dirSelectDir_Change()

' Input: None
' Process: Changes the file list contents to reflect the directory path
' Output: None

    On Error GoTo errorhandler ' Catching any errors in the following code
    filSelectFile.Path = dirSelectDir.Path ' Changing the list box
    Exit Sub
errorhandler: ' An error was found
    MsgBox "Error!  Path does not exist!", 16, "Error!"
End Sub

Private Sub drvSelectDrive_Change()

' Input: None
' Process: Changes the directory path to reflect the currently selected drive
' Output: None

    On Error GoTo errorhandler ' Catching any errors in the following code
    dirSelectDir.Path = drvSelectDrive.Drive ' Changing the path
    Exit Sub
errorhandler: ' An error was found
    MsgBox "Error!  Insert a disk and try again!", 16, "Error!"
End Sub

Private Sub filSelectFile_DblClick()

' Input: None
' Process: A file name has been selected and double-clicked.  This
'          simply calls the cmdOK_Click sub
' Output: None

    Call cmdOK_Click ' Calling the sub
End Sub
