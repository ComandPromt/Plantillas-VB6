VERSION 5.00
Begin VB.Form IconXTract 
   Caption         =   "IconXTract"
   ClientHeight    =   2220
   ClientLeft      =   3750
   ClientTop       =   1920
   ClientWidth     =   2100
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2220
   ScaleWidth      =   2100
   Begin VB.CommandButton Command1 
      Caption         =   "Show Small Icon"
      Height          =   300
      Left            =   312
      TabIndex        =   2
      Top             =   204
      Width           =   1476
   End
   Begin VB.PictureBox Picture1 
      Height          =   732
      Left            =   696
      ScaleHeight     =   675
      ScaleWidth      =   705
      TabIndex        =   1
      Top             =   1128
      Width           =   768
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Large Icon"
      Height          =   300
      Left            =   300
      TabIndex        =   0
      Top             =   660
      Width           =   1476
   End
End
Attribute VB_Name = "IconXTract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Sample VB4/32-bit code to retrieve the regular (32x32) and
'small (16x16) icons from an .EXE file without starting the program.
'Extraction techniques using ExtractIcon only return the 32x32 icon.
'Note: If the .EXE does not include a small icon, the regular icon will be
'produced reduced to 16x16, making the function appear to have worked.
'This sample is hard-coded to look at Explorer.exe, which does have both
'icons.
'Developed by Don Bradner with the assistance of Karl Peterson when a
'particularly nasty GPF wouldn't go away.  Feedback welcome to the Visual
'Basic Programmer's Journal forum on Compuserve (GO VBPJFORUM), in the
'32-bit section.

Option Explicit
Private Const MAX_PATH = 260
Private Const SHGFI_ICON = &H100
Private Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Private Const SHGFI_LARGEICON = &H0                      '  get large icon
Private Const SHGFI_SMALLICON = &H1                      '  get small icon
Private Const ILD_TRANSPARENT = &H1

Private Type SHFILEINFO 'Structure used by SHGetFileInfo
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal flags&) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private shinfo As SHFILEINFO
Private WinPath As String
Private xPixels As Integer
Private yPixels As Integer

Private Sub Command1_Click()
   Dim himl As Long
   Dim lpzxExeName As String '.EXE file name to get icon from
   Dim nRet As Long
   Dim picLeft As Long
   Dim picTop As Long

   lpzxExeName = WinPath & "\explorer.exe" 'Use any other executable that might contain a small icon
   himl = SHGetFileInfo(lpzxExeName, 0&, shinfo, Len(shinfo), SHGFI_SYSICONINDEX Or SHGFI_SMALLICON)
   
   '----------------------------------------------------
   'set the picture box up to receive the icon, centered
   '----------------------------------------------------
   picLeft = (Picture1.ScaleWidth / xPixels) / 2 - 8
   picTop = (Picture1.ScaleHeight / yPixels) / 2 - 8
   Picture1.Picture = LoadPicture() 'Clear any existing image
   Picture1.AutoRedraw = True
   nRet = ImageList_Draw(himl, shinfo.iIcon, Picture1.hDC, picLeft, picTop, ILD_TRANSPARENT)
   Picture1.Refresh
End Sub

Private Sub Command2_Click()
   Dim himl As Long
   Dim lpzxExeName As String '.EXE file name to get icon from
   Dim nRet As Long
   Dim picLeft As Long
   Dim picTop As Long
   
   lpzxExeName = WinPath & "\explorer.exe"
   himl = SHGetFileInfo(lpzxExeName, 0&, shinfo, Len(shinfo), SHGFI_SYSICONINDEX Or SHGFI_LARGEICON)
     
   '----------------------------------------------------
   'set the picture box up to receive the icon, centered
   '----------------------------------------------------
 
   picLeft = (Picture1.ScaleWidth / xPixels) / 2 - 16
   picTop = (Picture1.ScaleHeight / yPixels) / 2 - 16
   Picture1.Picture = LoadPicture()
   Picture1.AutoRedraw = True
   nRet = ImageList_Draw(himl, shinfo.iIcon, Picture1.hDC, picLeft, picTop, ILD_TRANSPARENT)
   Picture1.Refresh
End Sub
 

Private Sub Form_Load()
   Dim Buffer As String
   Dim nRet As Long
   Buffer = Space(MAX_PATH)
   nRet = GetWindowsDirectory(Buffer, Len(Buffer))
   WinPath = Left(Buffer, nRet)
   xPixels = Screen.TwipsPerPixelX
   yPixels = Screen.TwipsPerPixelY
End Sub


