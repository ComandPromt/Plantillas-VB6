Attribute VB_Name = "GetSysIcon"
Option Explicit
'Get Icons from System Image List
'Got This from someone; can't remember who
' =================================================================================
' Declares and types
' =================================================================================
Private Const MAX_PATH = 260
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, ByVal dwAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Enum EShellGetFileInfoConstants
     SHGFI_ICON = &H100                       ' // get icon
     SHGFI_DISPLAYNAME = &H200                ' // get display name
     SHGFI_TYPENAME = &H400                   ' // get type name
     SHGFI_ATTRIBUTES = &H800                 ' // get attributes
     SHGFI_ICONLOCATION = &H1000              ' // get icon location
     SHGFI_EXETYPE = &H2000                   ' // return exe type
     SHGFI_SYSICONINDEX = &H4000              ' // get system icon index
     SHGFI_LINKOVERLAY = &H8000               ' // put a link overlay on icon
     SHGFI_SELECTED = &H10000                 ' // show icon in selected state
     SHGFI_ATTR_SPECIFIED = &H20000           ' // get only specified attributes
     SHGFI_LARGEICON = &H0                    ' // get large icon
     SHGFI_SMALLICON = &H1                    ' // get small icon
     SHGFI_OPENICON = &H2                     ' // get open icon
     SHGFI_SHELLICONSIZE = &H4                ' // get shell size icon
     SHGFI_PIDL = &H8                         ' // pszPath is a pidl
     SHGFI_USEFILEATTRIBUTES = &H10           ' // use passed dwFileAttribute
End Enum
Private Type PictDesc
  cbSizeofStruct As Long
  picType As Long
  hImage As Long
  xExt As Long
  yExt As Long
End Type
Private Type Guid
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal x&, ByVal y&, ByVal flags&) As Long

' =================================================================================
' Interface
' =================================================================================
Public Enum EGetIconTypeConstants
    egitSmallIcon = 1
    egitLargeIcon = 2
End Enum
Private Const ILD_TRANSPARENT = &H1       'Display transparent

Public Function GetIcon( _
        ByVal sFIle As String, _
        Optional ByVal EIconType As EGetIconTypeConstants = egitLargeIcon _
    ) As Object
Dim lR As Long
Dim r As Long
Dim hIcon As Long
Dim iIcon As Long
Dim tSHI As SHFILEINFO
Dim lFlags As Long
    
    ' Prepare flags for SHGetFileInfo to get the icon:
    If (EIconType = egitLargeIcon) Then
        lFlags = SHGFI_ICON Or SHGFI_LARGEICON
    Else
        lFlags = SHGFI_ICON Or SHGFI_SMALLICON
    End If
    lFlags = lFlags And Not SHGFI_LINKOVERLAY
    lFlags = lFlags And Not SHGFI_OPENICON
    lFlags = lFlags And Not SHGFI_SELECTED
    ' Call to get icon:
    lR = SHGetFileInfo(sFIle, 0&, tSHI, Len(tSHI), lFlags)
    If (lR <> 0) Then
        ' If we succeeded, the hIcon member will be filled in:
        hIcon = tSHI.hIcon
        iIcon = tSHI.iIcon
        ' If we have an icon, convert it to a VB picture and return it:
        If (hIcon <> 0) Then
            Set GetIcon = IconToPicture(hIcon)
        End If
        ' Free resouce:
        DeleteObject hIcon
    End If
    
End Function
Public Function IconToPicture(ByVal hIcon As Long) As IPicture
    
    If hIcon = 0 Then Exit Function
        
    ' This is all magic if you ask me:
    Dim NewPic As Picture, PicConv As PictDesc, IGuid As Guid
    
    PicConv.cbSizeofStruct = Len(PicConv)
    PicConv.picType = vbPicTypeIcon
    PicConv.hImage = hIcon
    
    'IGuid.Data1 = &H20400
    'IGuid.Data4(0) = &HC0
    'IGuid.Data4(7) = &H46
    ' Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    OleCreatePictureIndirect PicConv, IGuid, True, NewPic
    
    Set IconToPicture = NewPic
    
End Function
Public Function GetFileTypeName( _
        ByVal sFIle As String _
    ) As String
Dim lR As Long
Dim tSHI As SHFILEINFO
Dim iPos As Long

    lR = SHGetFileInfo(sFIle, 0&, tSHI, Len(tSHI), SHGFI_TYPENAME)
    If (lR <> 0) Then
        iPos = InStr(tSHI.szTypeName, Chr$(0))
        If (iPos = 0) Then
            GetFileTypeName = tSHI.szTypeName
        ElseIf (iPos > 1) Then
            GetFileTypeName = Left$(tSHI.szTypeName, (iPos - 1))
        Else
            GetFileTypeName = ""
        End If
    End If
    
End Function






