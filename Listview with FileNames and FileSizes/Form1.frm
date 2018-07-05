VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdReport 
      Caption         =   "&Report View"
      Height          =   525
      Left            =   2460
      TabIndex        =   6
      Top             =   4470
      Width           =   1275
   End
   Begin VB.CommandButton CmdSmall 
      Caption         =   "&Small Icons"
      Height          =   525
      Left            =   1230
      TabIndex        =   5
      Top             =   4470
      Width           =   1185
   End
   Begin VB.CommandButton CmdLarge 
      Caption         =   "&Large Icons"
      Height          =   525
      Left            =   60
      TabIndex        =   4
      Top             =   4470
      Width           =   1125
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   30
      TabIndex        =   2
      Top             =   690
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6588
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "Name"
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   "Size"
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   "CreatedOn"
         Object.Tag             =   ""
         Text            =   "Creation Date"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   900
      TabIndex        =   0
      Top             =   30
      Width           =   4635
   End
   Begin VB.Label LBLNewDir 
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   390
      Width           =   6465
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   6150
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0000
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0352
            Key             =   "File"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5550
      Top             =   -60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":06A4
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":12F6
            Key             =   "File"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Drive:"
      Height          =   225
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@================================================================================
' do not distribute this code in uncompiled form. Also do not display it on a web
' page. Its contents where created for vbplanet.dyndns.org and can only be
' displayed there
'=================================================================================
Option Compare Text
Option Explicit
' Constants
Const vbFindAll = 63 ' All attribs (for dir())
'@==============================================
' cmdLarge_Click:
'   Show Large Icons
'===============================================
Private Sub CmdLarge_Click()
 ListView1.View = lvwIcon
End Sub
'@==============================================
' cmdReport_Click:
'   Show report view
'===============================================
Private Sub CmdReport_Click()
 ListView1.View = lvwReport
End Sub
'@==============================================
' cmdSmall_Click:
'   Show small Icons
'===============================================
Private Sub CmdSmall_Click()
 ListView1.View = lvwSmallIcon
End Sub

'@=======================================
' Drive1_Change:
' update listview
'=========================================
Private Sub Drive1_Change()
 Listview1_ChangeDir (Mid(Drive1.Drive, 1, 1) + ":")
End Sub
'@============================================
' Listview1_ChangeDir
'   Changes listview to display <NEWDIR>
'==============================================
Private Sub Listview1_ChangeDir(ByVal NewDir As String)
 Dim TXT As String ' General
' show newdir
 LBLNewDir = NewDir
 ' clear listitems
 ListView1.ListItems.Clear
 ' get dirs
 TXT = Dir(NewDir + "\", vbFindAll)
 ' add rest of files
 While Not TXT = ""
  ' if its a directory
  If GetAttr(NewDir + "\" + TXT) And vbDirectory Then
   ListView1.ListItems.Add , , TXT, "Folder", "Folder"
   ' add size property
   ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "Folder"
   ' add Created On
    If Not TXT = "." And Not TXT = ".." Then
     ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = _
     CStr(FileDateTime(NewDir + "\" + TXT))
    End If
  End If
  TXT = Dir
 Wend
  ' get files
 TXT = Dir(NewDir + "\", vbFindAll)
 ' add rest of files
 While Not TXT = ""
  ' if its a directory
  If Not GetAttr(NewDir + "\" + TXT) And vbDirectory Then
   ListView1.ListItems.Add , , TXT, "File", "File"
   ' add size property
   ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = _
   CStr(FileLen(NewDir + "\" + TXT))
   ' add Created On
   ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = _
   CStr(FileDateTime(NewDir + "\" + TXT))
  End If
  TXT = Dir
 Wend
 ListView1.Arrange = lvwAutoLeft
End Sub
'@==============================================
' Form_Load: here we get files for the first time
'================================================
Private Sub Form_Load()
 Drive1_Change
End Sub
'@===============================================
' ListView1_DblClick:
'   Open directory
'================================================
Private Sub ListView1_DblClick()
 Dim A As Long ' general use
 Dim TXT As String ' General
 Dim Dirs() As String ' for holding dirs
 ' if . we exit
 If ListView1.SelectedItem.Text = "." Then Exit Sub
' if we need to go back a dir
 If ListView1.SelectedItem.Text = ".." Then
  ' get path
   TXT = LBLNewDir
   ' get dirs
   Dirs = Split(TXT, "\")
   ' remove first one
   TXT = Dirs(0)
    For A = 1 To UBound(Dirs) - 1
     TXT = TXT + "\" + Dirs(A)
     Next A
   If GetAttr(TXT) And vbDirectory Then Listview1_ChangeDir (TXT)
  Else ' we just go into dir
  ' get fullpath
  TXT = LBLNewDir + "\" + ListView1.SelectedItem.Text
  ' if directory we open
  If GetAttr(TXT) And vbDirectory Then Listview1_ChangeDir (TXT)
 End If ' end if .. or not
End Sub
