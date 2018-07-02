VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Shfileop sample from BlackBeltVB.com"
   ClientHeight    =   2685
   ClientLeft      =   330
   ClientTop       =   1515
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2685
   ScaleWidth      =   7185
   Begin VB.CommandButton Command1 
      Caption         =   "&Just Do It!"
      Default         =   -1  'True
      Height          =   495
      Left            =   1740
      TabIndex        =   8
      Top             =   2100
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   720
      TabIndex        =   6
      Text            =   "SHFileOp Sample Move to the Recycle Bin"
      Top             =   1620
      Width           =   3435
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Simple progress box (no filenames)"
      Height          =   195
      Index           =   3
      Left            =   420
      TabIndex        =   5
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Display progress box"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   780
      Value           =   1  'Checked
      Width           =   1755
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Confirm operations"
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   1635
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Permanent delete"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   3660
      TabIndex        =   1
      Text            =   "c:\windows\temp\*.*"
      Top             =   60
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"shfileop.frx":0000
      Height          =   1875
      Index           =   2
      Left            =   4440
      TabIndex        =   9
      Top             =   600
      Width           =   2595
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Progress Box Title"
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Top             =   1380
      Width           =   3435
   End
   Begin VB.Label Label1 
      Caption         =   "Enter directory or file(s) to move to the recycle bin."
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Shfileop sample from BlackBeltVB.com
' http://blackbeltvb.com
'
' Written by Matt Hart
' Copyright 1999 by Matt Hart
'
' This software is FREEWARE. You may use it as you see fit for
' your own projects but you may not re-sell the original or the
' source code. Do not copy this sample to a collection, such as
' a CD-ROM archive. You may link directly to the original sample
' using "http://blackbeltvb.com/shfileop.htm"
'
' No warranty express or implied, is given as to the use of this
' program. Use at your own risk.
'
' This program shows how to implement the SHFileOperation API.
' You can use it to delete, move, or copy multiple or single files,
' and it can send files to the recycle bin.

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

' // Shell File Operations

Const FO_MOVE = &H1
Const FO_COPY = &H2
Const FO_DELETE = &H3
Const FO_RENAME = &H4
Const FOF_MULTIDESTFILES = &H1
Const FOF_CONFIRMMOUSE = &H2
Const FOF_SILENT = &H4                      '  don't create progress/report
Const FOF_RENAMEONCOLLISION = &H8
Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.
Const FOF_WANTMAPPINGHANDLE = &H20          '  Fill in SHFILEOPSTRUCT.hNameMappings
                                      '  Must be freed using SHFreeNameMappings
Const FOF_ALLOWUNDO = &H40
Const FOF_FILESONLY = &H80                  '  on *.*, do only files
Const FOF_SIMPLEPROGRESS = &H100            '  means don't show names of files
Const FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed dirs

Const PO_DELETE = &H13           '  printer is being deleted
Const PO_RENAME = &H14           '  printer is being renamed
Const PO_PORTCHANGE = &H20       '  port this printer connected to is being changed
                                '  if this id is set, the strings received by
                                '  the copyhook are a doubly-null terminated
                                '  list of strings.  The first is the printer
                                '  name and the second is the printer port.
Const PO_REN_PORT = &H34         '  PO_RENAME and PO_PORTCHANGE at same time.

Private Sub Check1_Click(Index As Integer)
    If Check1(2).Value = 0 Then
        Check1(3).Enabled = False
        Text2.Enabled = False
    Else
        Check1(3).Enabled = True
        Text2.Enabled = True
    End If
End Sub

'  no POF_ flags currently defined

'  implicit parameters are:
'       if pFrom or pTo are unqualified names the current directories are
'       taken from the global current drive/directory settings managed
'       by Get/SetCurrentDrive/Directory
'
'       the global confirmation settings

Private Sub Command1_Click()
    Dim lResult As Long, SHF As SHFILEOPSTRUCT
    SHF.hwnd = hwnd
    SHF.wFunc = FO_DELETE
    SHF.pFrom = Text1.Text
    SHF.fFlags = FOF_FILESONLY
    If Check1(0).Value = 0 Then SHF.fFlags = SHF.fFlags + FOF_ALLOWUNDO
    If Check1(1).Value = 0 Then SHF.fFlags = SHF.fFlags + FOF_NOCONFIRMATION
    If Check1(2).Value = 0 Then
        SHF.fFlags = SHF.fFlags + FOF_SILENT
    Else
        If Check1(3).Value Then SHF.fFlags = SHF.fFlags + FOF_SIMPLEPROGRESS
        SHF.lpszProgressTitle = Text2.Text
    End If
    lResult = SHFileOperation(SHF)
    If lResult Then
        MsgBox "Error occurred!"
    End If
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub
