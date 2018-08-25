Attribute VB_Name = "modIcon"
Option Explicit
Private Const MAX_PATH = 260
Private Const ILD_TRANSPARENT = &H1
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal FLAGS&) As Long
Function GetIcon(FileName As String, pIcon As PictureBox) As String
Dim hSIcon As Long
Dim ShInfo As SHFILEINFO
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
GetIcon = ShInfo.szDisplayName
Set pIcon.Picture = Nothing
pIcon.Cls
ImageList_Draw hSIcon, ShInfo.iIcon, pIcon.hDC, 0, 0, ILD_TRANSPARENT
pIcon.Refresh
End Function
