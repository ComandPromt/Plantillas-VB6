VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CICON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const MAX_PATH = 260
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000  'system icon index
Private Const SHGFI_LARGEICON = &H0  'large icon
Private Const SHGFI_SMALLICON = &H1  'small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const ILD_TRANSPARENT = &H1  'display transparent
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or _
                                 SHGFI_SHELLICONSIZE Or _
                                 SHGFI_SYSICONINDEX Or _
                                 SHGFI_DISPLAYNAME Or _
                                 SHGFI_EXETYPE
Private Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type
Private Declare Function SHGetFileInfo Lib _
   "shell32.dll" Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" _
   (ByVal himl As Long, ByVal i As Long, _
    ByVal hDCDest As Long, ByVal x As Long, _
    ByVal y As Long, ByVal flags As Long) As Long
Private shinfo As SHFILEINFO
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@saic.com]
' Date: September,19 2000 @ 11:53:13
'------------------------------------------------------------
Public Function ExtractIconToHDC(hdc As Long, fNAME As String) As Boolean
    On Error GoTo ErrorExtractIconToHDC
    Dim hImgSmall As Long   'the handle to the system image list
   'get the system icon associated with that file
    hImgSmall& = SHGetFileInfo(fNAME, 0&, _
        shinfo, Len(shinfo), _
        BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    'draw the associated icon into the picturebox
   Call ImageList_Draw(hImgSmall&, shinfo.iIcon, _
                       hdc, 0, 0, ILD_TRANSPARENT)
    Exit Function
ErrorExtractIconToHDC:
    MsgBox Err & ":Error in ExtractIconToHDC.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Function
End Function


