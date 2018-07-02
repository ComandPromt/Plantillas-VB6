VERSION 5.00
Begin VB.Form frmGetPath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extract"
   ClientHeight    =   1545
   ClientLeft      =   3960
   ClientTop       =   2430
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetPath.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Double Click to select the folder to extract to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   5415
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   5055
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.TextBox txtSaveAs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   3795
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Extract to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtFolderName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   660
         Width           =   3795
      End
      Begin VB.OptionButton optBrowse 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   255
      End
      Begin VB.OptionButton optSameFolderAsCab 
         Caption         =   "Same folder as cab file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton cmdBrowseForFolder 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   390
         TabIndex        =   2
         Top             =   660
         Width           =   855
      End
   End
   Begin VB.Label lblSaveAs 
      Caption         =   "Save File As:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "frmGetPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnCanceled     As Boolean
Private mstrPath         As String
Private mstrOriginalPath As String

Public SingleFile        As Boolean
Public FileName          As String
Public Property Get Canceled() As Boolean
    Canceled = mblnCanceled
End Property
Public Property Let Path(Value As String)
    mstrPath = Value
    mstrOriginalPath = Value
End Property
Public Property Get Path() As String
    Path = mstrPath
End Property
Private Sub cmdBrowseForFolder_Click()

    optBrowse.Value = True

End Sub
Private Sub cmdCancel_Click()
    mblnCanceled = True
    Unload Me
End Sub
Private Sub cmdExtract_Click()
    mblnCanceled = False
    
    If optSameFolderAsCab Then
        mstrPath = mstrOriginalPath
    Else
        mstrPath = txtFolderName
    End If
    
    FileName = txtSaveAs
    Unload Me
End Sub
Private Sub Dir1_Change()
    
    txtFolderName = Dir1.Path

End Sub
Private Sub Drive1_Change()

    Dir1.Path = Drive1.Drive
    
End Sub
Private Sub Form_Load()
    '
    ' Initialize values
    '
    Drive1.Drive = Left$(App.Path, 3)
    
    txtFolderName = mstrPath
    txtSaveAs = FileName
    optSameFolderAsCab.Value = True
    mblnCanceled = False
    lblSaveAs.Visible = SingleFile
    txtSaveAs.Visible = SingleFile
End Sub
Private Sub optBrowse_Click()

    Dir1.Visible = optBrowse.Value
    Drive1.Visible = optBrowse.Value
    Me.Height = 4605
    
End Sub
Private Sub optSameFolderAsCab_Click()

    Dir1.Visible = optBrowse.Value
    Drive1.Visible = optBrowse.Value
    Me.Height = 1950

End Sub
Private Sub txtFolderName_Change()
    optBrowse.Value = True
End Sub



