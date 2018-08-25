VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLarge 
      Caption         =   "Extract Large"
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Top             =   1140
      Width           =   1515
   End
   Begin VB.PictureBox picBuffer 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   2700
      ScaleHeight     =   360
      ScaleWidth      =   1560
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox picViewIcon 
      Height          =   1215
      Left            =   180
      ScaleHeight     =   1155
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   1740
      Width           =   2415
   End
   Begin VB.TextBox txtFilename 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   900
      Width           =   2415
   End
   Begin VB.CommandButton cmdSmall 
      Caption         =   "Extract Small"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   1515
   End
   Begin MSComctlLib.ImageList imgIconList 
      Left            =   3240
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Type a filename below of which you wish to extract the icon from:"
      Height          =   495
      Left            =   180
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long

Private Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Dim FileInfo As typSHFILEINFO

Private Sub cmdLarge_Click()
    Dim r As Integer
    r = ExtractIcon(txtFilename.Text, imgIconList, picBuffer, 32)
    
    If r = 0 Then
        MsgBox "Path not found!"
    Else
        picViewIcon.Picture = imgIconList.ListImages(r).Picture
    End If
End Sub

Private Sub cmdSmall_Click()
    Dim r As Integer
    r = ExtractIcon(txtFilename.Text, imgIconList, picBuffer, 16)
    
    If r = 0 Then
        MsgBox "Path not found!"
    Else
        picViewIcon.Picture = imgIconList.ListImages(r).Picture
    End If
End Sub

Private Function ExtractIcon(FileName As String, AddtoImageList As ImageList, PictureBox As PictureBox, PixelsXY As Integer) As Long
    Dim SmallIcon As Long
    Dim NewImage As ListImage
    Dim IconIndex As Integer
    
    If PixelsXY = 16 Then
        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
    Else
        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
    End If
    
    If SmallIcon <> 0 Then
      With PictureBox
        .Height = 15 * PixelsXY
        .Width = 15 * PixelsXY
        .ScaleHeight = 15 * PixelsXY
        .ScaleWidth = 15 * PixelsXY
        .Picture = LoadPicture("")
        .AutoRedraw = True
        
        SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, PictureBox.hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      
      IconIndex = AddtoImageList.ListImages.Count + 1
      Set NewImage = AddtoImageList.ListImages.Add(IconIndex, , PictureBox.Image)
      ExtractIcon = IconIndex
    End If
End Function
