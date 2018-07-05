VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExtrIconsTest 
   Caption         =   "Test Icon Extraction"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbStyle 
      Height          =   315
      ItemData        =   "frmExtrIconsTest.frx":0000
      Left            =   2880
      List            =   "frmExtrIconsTest.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3900
      Width           =   2055
   End
   Begin VB.PictureBox pic16 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1920
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   3780
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic32 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1380
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   3780
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   5100
      TabIndex        =   3
      Top             =   3840
      Width           =   1635
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   180
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3015
      Left            =   180
      TabIndex        =   2
      Top             =   660
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "File"
         Text            =   "File"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Path"
         Text            =   "Path"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   780
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   180
      Width           =   5895
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   780
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Path"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   330
   End
End
Attribute VB_Name = "frmExtrIconsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------
'Module     : frmExtrIconsTest
'Description: Test program to demo icon extraction
'Release    : 2001 VB6 SP4
'Copyright  : © T De Lange, 2000
'E-mail     : tomdl@attglobal.net
'----------------------------------------------------------------
'This project demonstrates how to extract icons associated with
'files into an imagelist and displaying them in a listview with
'the filenames.
'The SHGetFileInfo function of the shell32.dll library is used,
'which makes the job much easier than before. The ImageList_Draw
'function in comctl32.dll is used to draw the icon in a picture box,
'from where it is placed into the image list.
'Watch out for the following:
'a) Image list can hold only approx 400 icons, so you will have
'   to remove duplicate images for files other than exe's
'b) Remember to set the lvw's mask color to the appropriate
'   system color, usually buttonface.
'----------------------------------------------------------------
'Credits:
'Peter Meier, Planet Source Code for
'the technique as used in his 'DelRecent' posting
'----------------------------------------------------------------
Option Base 0
Option Explicit
DefLng A-N, P-Z
DefBool O

'Icon Sizes in pixels
Private Const LARGE_ICON As Integer = 32
Private Const SMALL_ICON As Integer = 16
Private Const MAX_PATH = 260

Private Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type


'----------------------------------------------------------
'Functions & Procedures
'----------------------------------------------------------
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal x&, ByVal y&, ByVal flags&) As Long


'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Private ShInfo As SHFILEINFO

Sub FillLvwWithFiles(ByVal Path As String)
'-------------------------------------------
'Scan the selected folder for files
'and add then to the listview
'-------------------------------------------
Dim Item As ListItem
Dim s As String

Path = CheckPath(Path)    'Add '\' to end if not present
s = Dir(Path, vbNormal)
Do While s <> ""
  Set Item = lvw.ListItems.Add()
  Item.Key = Path & s
  'Item.SmallIcon = "Folder"
  Item.Text = s
  Item.SubItems(1) = Path
  s = Dir
Loop

End Sub
Private Function CheckPath(ByVal Path As String) As String
'--------------------------------------------------
'Checks if path ends with "\". If not, add it.
'--------------------------------------------------
If Right(Path, 1) <> "\" Then
  CheckPath = Path & "\"
Else
  CheckPath = Path
End If

End Function

Private Sub cmbStyle_Click()
lvw.View = cmbStyle.ListIndex

End Sub


Private Sub cmdShow_Click()
'-------------------------------------------
'Load the files into the listview
'-------------------------------------------
Dim Path As String

Initialise
Path = txtPath.Text
FillLvwWithFiles Path
GetAllIcons
ShowIcons

End Sub

Private Sub Initialise()
'-----------------------------------------------
'Initialise the controls
'-----------------------------------------------
On Local Error Resume Next

'Break the link to iml lists
lvw.ListItems.Clear
lvw.Icons = Nothing
lvw.SmallIcons = Nothing

'Clear the image lists
iml32.ListImages.Clear
iml16.ListImages.Clear

End Sub

Private Sub GetAllIcons()
'--------------------------------------------------
'Extract all icons
'--------------------------------------------------
Dim Item As ListItem
Dim FileName As String

On Local Error Resume Next
For Each Item In lvw.ListItems
  FileName = Item.SubItems(1) & Item.Text
  GetIcon FileName, Item.Index
Next

End Sub

Private Function GetIcon(FileName As String, Index As Long) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long


'Get a handle to the small icon
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then
  'Large Icon
  With pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hLIcon, ShInfo.iIcon, pic32.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  'Small Icon
  With pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hSIcon, ShInfo.iIcon, pic16.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  Set imgObj = iml32.ListImages.Add(Index, , pic32.Image)
  Set imgObj = iml16.ListImages.Add(Index, , pic16.Image)
End If
End Function
Private Sub ShowIcons()
'-----------------------------------------
'Show the icons in the lvw
'-----------------------------------------
On Error Resume Next

Dim Item As ListItem
With lvw
  '.ListItems.Clear
  .Icons = iml32        'Large
  .SmallIcons = iml16   'Small
  For Each Item In .ListItems
    Item.Icon = Item.Index
    Item.SmallIcon = Item.Index
  Next
End With

End Sub


Private Sub Form_Load()
'---------------------------------------------
'Once off initialisations
'---------------------------------------------

'Size the picture boxes containing the icons
pic16.Width = (SMALL_ICON) * Screen.TwipsPerPixelX
pic16.Height = (SMALL_ICON) * Screen.TwipsPerPixelY
pic32.Width = LARGE_ICON * Screen.TwipsPerPixelX
pic32.Height = LARGE_ICON * Screen.TwipsPerPixelY

cmbStyle.ListIndex = lvw.View

End Sub


