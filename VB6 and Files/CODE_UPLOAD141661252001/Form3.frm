VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3330
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7905
   LinkTopic       =   "Form3"
   ScaleHeight     =   3330
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6435
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   405
      Left            =   6420
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2265
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.ListBox lstSections 
      Height          =   2010
      Left            =   3600
      TabIndex        =   1
      Top             =   675
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.ListBox lstKeys 
      Height          =   2010
      Left            =   705
      TabIndex        =   0
      Top             =   675
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Label Label2 
      Caption         =   "Sections On File"
      Height          =   225
      Left            =   3600
      TabIndex        =   4
      Top             =   345
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "Keys On File"
      Height          =   225
      Left            =   720
      TabIndex        =   3
      Top             =   345
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Menu mnuAvailable 
      Caption         =   "&Available Options"
      Begin VB.Menu mnuListKeys 
         Caption         =   "Show All Keys of MySystem.ini"
      End
      Begin VB.Menu mnuListSections 
         Caption         =   "Show All Sections of MySystem.ini"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Set My Color Preference for Form2, Form3"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A program by Legrev3@aol.com
'Submitted for downloading Dec 6, 2000
'Demonstrates usage of module ReadWrite.bas for maintaining .ini file

Option Explicit

Private Sub cmdExit_Click()
    Dim i As Integer
    On Error Resume Next
    For i = 0 To Forms.Count - 1
        Unload Forms(i)
    Next i
    End
End Sub



Private Sub Form_Load()
    Me.Caption = "Welcome " & strLoginName & "!"
    strSection = "User Preferences"
    strColor = ReadFromFile(strSection, strLoginName)
    If strColor = "" Then
        lngColor = &H8000000F            'no user color preferences use system color
    Else
        lngColor = CLng(strColor)
    End If
    
    Me.BackColor = lngColor
End Sub

Private Sub mnuColor_Click()
    
    CommonDialog1.CancelError = True
    On Error GoTo Cancelled:
    CommonDialog1.Flags = cdlCCRGBInit
    CommonDialog1.ShowColor
    Form2.BackColor = CommonDialog1.Color
    Form3.BackColor = CommonDialog1.Color
    Form3.Label1.BackColor = CommonDialog1.Color
    Form3.Label2.BackColor = CommonDialog1.Color
    Form3.cmdExit.BackColor = CommonDialog1.Color
    
   'save user preferences to file
    strSection = "User Preferences"
    lngColor = Me.BackColor
    strColor = CStr(lngColor)
    lngRetVal = WriteToFile(strSection, strLoginName, strColor)
Cancelled:
End Sub

Public Sub mnuListKeys_Click()
    Dim strAllKeys As String
    Dim strOneKey As String
    Dim intPos1 As Long
    Dim intPos2 As Long
    
    lstKeys.Visible = True
    lstKeys.Clear
    
    With Label1
        .Visible = True
        .BackColor = Me.BackColor
    End With
    
    With cmdExit
        .Visible = True
        .BackColor = Me.BackColor
    End With
    
    'read all keys (usernames) from password section of file
    strAllKeys = String$(BUFF_SIZ, Space$(1))
    strSection = "Password Section"
    lngRetVal = GetPrivateProfileStringKeys(strSection, 0, "", strAllKeys, BUFF_SIZ - 1, strMySystemFile)
    
    strAllKeys = Trim$(strAllKeys)
    intPos1 = InStr(strAllKeys, Chr(0))
    intPos2 = 1
    
    'add each to list box
    While intPos2 <> Len(strAllKeys)
        intPos1 = InStr(intPos2, strAllKeys, Chr(0))
        strOneKey = Mid$(strAllKeys, intPos2, intPos1 - intPos2)
        lstKeys.AddItem strOneKey
        intPos2 = intPos1 + 1
    Wend
End Sub

Public Sub mnuListSections_Click()
    Dim strAllSections As String
    Dim strOneSection As String
    Dim intPos1 As Long
    Dim intPos2 As Long
    
    lstSections.Visible = True
    lstSections.Clear
    
    With Label2
        .Visible = True
        .BackColor = Me.BackColor
    End With
    
    With cmdExit
        .Visible = True
        .BackColor = Me.BackColor
    End With
    
    'read all section names from file
    strAllSections = String$(BUFF_SIZ, Space$(1))
    strSection = "Password Section"
    lngRetVal = GetPrivateProfileStringSections(0, 0, "", strAllSections, BUFF_SIZ, strMySystemFile)
    
    strAllSections = Trim$(strAllSections)

    intPos1 = InStr(strAllSections, Chr(0))
    intPos2 = 1
    
    'add each to listbox
    While intPos2 <> Len(strAllSections)
        intPos1 = InStr(intPos2, strAllSections, Chr(0))
        strOneSection = Mid$(strAllSections, intPos2, intPos1 - intPos2)
        lstSections.AddItem strOneSection
        intPos2 = intPos1 + 1
    Wend
End Sub

