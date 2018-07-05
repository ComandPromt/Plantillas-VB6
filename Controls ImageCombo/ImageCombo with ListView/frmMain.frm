VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Image Combo/ListView"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSelFile 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   3000
      Width           =   4695
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
      ImageList       =   "ilsIcons16"
   End
   Begin MSComctlLib.ImageList ilsIcons16 
      Left            =   3360
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "MyComputer"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsIcons32 
      Left            =   2520
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3201
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ilsIcons32"
      SmallIcons      =   "ilsIcons16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Size"
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Type"
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Date"
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1680
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3904
            Key             =   "Up One Level"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A16
            Key             =   "Clsdfold"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D30
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E42
            Key             =   "View Details"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F54
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4066
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4178
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   4320
      TabIndex        =   5
      Top             =   600
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up One Level"
            Object.ToolTipText     =   "Up One Level"
            ImageKey        =   "Up One Level"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New Folder"
            Object.ToolTipText     =   "New Folder"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Details"
            Object.ToolTipText     =   "View Details"
            ImageKey        =   "View Details"
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View List"
            Object.ToolTipText     =   "View List"
            ImageKey        =   "View List"
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Large Icons"
            Object.ToolTipText     =   "View Large Icons"
            ImageKey        =   "View Large Icons"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   4200
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   840
      TabIndex        =   7
      Top             =   3120
      Width           =   45
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject
Dim d, dc, s, n, t
Dim fs, f, f1, fc, s1, nf

Dim bIsDrive As Boolean
Dim ci As ComboItem
Dim Children() As String
Dim ChildIndex As Variant

Dim imgX As ListImage
Dim itmFldr As ListItem
Dim itmX As ListItem
Dim iListCount As Integer
Dim iParentIndex As Integer
Dim iNumOfChildren As Integer
Dim sCheckChild As String
Dim sDir As String
Dim sItem As String
Dim sKey As String
Dim sKey1 As String
Dim sPath As String
Dim sText As String

Private Sub Dir1_Change()
File1.Path = Dir1.Path   ' Set file path.
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive   ' Set directory path.
End Sub

Private Sub Form_Load()
Dir1.Path = "C:\"
ListView1.View = lvwList
Toolbar1.Buttons("View List").Value = 1
GetDrives
'Add Desktop
ImageCombo1.ComboItems.Add 1, "c:\WINDOWS\Desktop", "Desktop", 4
'Add the Root "My Computer" to ImageCombo1
ImageCombo1.ComboItems.Add 2, "Root", "My Computer", 1, , 1
'Select "C Drive"
ImageCombo1.ComboItems("c:\").Selected = True
bIsDrive = True
ImageCombo1_Click
End Sub

Public Sub GetDrives()
    For i = 0 To Drive1.ListCount - 1
        d = Drive1.List(i)
        sText = d
        sKey = Left$(sText, 2) & "\"
        Set imgX = ilsIcons16.ListImages.Add(, sKey, GetIcon(sKey, egitSmallIcon))
        ImageCombo1.ComboItems.Add i + 1, sKey, sText, sKey, sKey, 2
    Next
    
End Sub

Private Sub ImageCombo1_Click()
    txtSelFile = ""
    On Error Resume Next
    If ImageCombo1.SelectedItem.Key = "c:\WINDOWS\Desktop" Then
        bIsDrive = False
        GetDesktop
        Exit Sub
    ElseIf ImageCombo1.SelectedItem.Key = "Root" Then
        'bIsDrive = False
        GetMyComputer
        Exit Sub
    End If
    Drive1.Drive = ImageCombo1.SelectedItem.Key
    Dir1.Path = ImageCombo1.SelectedItem.Key
    'Get selected Item
    sDrive = ImageCombo1.SelectedItem.Key
    iLength = Len(sDrive)
    'Store Index for AddChild
    iParentIndex = ImageCombo1.SelectedItem.Index
    If iLength > 3 Then 'Selected Item is a Folder
        bIsDrive = False
        ListView1.ListItems.Clear
        File1.Path = ImageCombo1.SelectedItem.Key
        ShowFolderList
        ShowFileList
    Else 'Selected Item is a Drive
        bIsDrive = True
        DeleteChild 'Delete previous children
        Set drv = fso.GetDrive(fso.GetDriveName(sDrive))
        If drv.IsReady Then
            f = drv.RootFolder
            'Convert to lower case to avoid UpOne error
            f = LCase(f)
            ListView1.ListItems.Clear
            ShowFolderList
            ShowFileList
        Else
            MsgBox "Drive " & drv & " Not Ready", 16
        End If
        ListView1.View = lvwList
        Toolbar1.Buttons("View List").Value = 1
        Toolbar1.Buttons("View Details").Value = 0
        Toolbar1.Buttons("View Large Icons").Value = 0
    End If
    
End Sub

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    sKey = ImageCombo1.SelectedItem.Key
    AddNewFolder sKey, NewString
    fso.DeleteFolder sKey & "New Folder"
End Sub

Private Sub ListView1_DblClick()

On Error Resume Next
    'Get the selected Item
  sItem = ListView1.SelectedItem.Key
  File1.Path = sItem
  sDir = ListView1.SelectedItem
  
  If Len(sItem) = 3 Then
    bIsDrive = True
    ImageCombo1.ComboItems(sItem).Selected = True
    ImageCombo1_Click
    'If it is a file select it
  ElseIf InStr(sItem, ".") > 0 Then
    Item.Selected = True
    txtSelFile = sItem
  Else 'If it is a Directory add it to the ImageCombo
    ListView1.ListItems.Clear
    Dir1.Path = sItem
    bIsDrive = False
    Screen.MousePointer = vbHourglass
    ShowFolderList
    ShowFileList     'Get the contents
    Screen.MousePointer = vbDefault
    'Check if already added
        If CheckChild(sItem) Then 'If added, Just select it
            ImageCombo1.ComboItems(sItem).Selected = True
        Else 'Add New Child
            AddChild sItem, sDir
        End If
    ListView1.View = lvwList
    Toolbar1.Buttons("View List").Value = 1
    Toolbar1.Buttons("View Details").Value = 0
    Toolbar1.Buttons("View Large Icons").Value = 0
  End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    sItem = ListView1.SelectedItem.Key
    If InStr(sItem, ".") > 0 Then
        Item.Selected = True
        txtSelFile = sItem
        Label1 = "Selected File:"
    Else
        Item.Selected = True
        txtSelFile = sItem
        Label1 = "Selected Folder:"
    End If
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.Key
        
        Case "View Large Icons"
            ListView1.View = lvwIcon
            Toolbar1.Buttons("View List").Value = 0
            Toolbar1.Buttons("View Details").Value = 0
        Case "Delete"
            'Add code for no file selected
            Dim Msg, Style, Title, Help, Ctxt, Response
            Msg = "Delete File?"   ' Define message.
            Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
            Title = "Verify Delete File"   ' Define title.
            Response = MsgBox(Msg, Style, Title)
            If Response = vbYes Then   ' User chose Yes.
               fso.DeleteFile txtSelFile
            Else   ' User chose No.
               Exit Sub
            End If
        Case "View List"
            Toolbar1.Buttons("View Details").Value = 0
            Toolbar1.Buttons("View Large Icons").Value = 0
            ListView1.View = lvwList
        Case "View Details"
            Toolbar1.Buttons("View List").Value = 0
            Toolbar1.Buttons("View Large Icons").Value = 0
            ListView1.View = lvwReport
            ListView1.Sorted = False
            folderspec = ImageCombo1.SelectedItem.Key
            GetlvReportData folderspec
        Case "New Folder"
            sPath = ImageCombo1.SelectedItem.Key
            AddNewFolder sPath, ""
            ImageCombo1_Click
            ListView1.SetFocus
            Set ListView1.SelectedItem = ListView1.ListItems(sPath & "New Folder")
            ListView1.SelectedItem.EnsureVisible
            ListView1.StartLabelEdit
        Case "Up One Level"
            txtSelFile = ""
            sKey = ImageCombo1.SelectedItem.Key
            sParentFolder = fso.GetParentFolderName(sKey)
            ListView1.View = lvwList
            Toolbar1.Buttons("View List").Value = 1
            Toolbar1.Buttons("View Details").Value = 0
            Toolbar1.Buttons("View Large Icons").Value = 0
            'code to adjust for Drive being selected
            If (Len(sParentFolder) = 3) Then
                'Parent is a Drive
                ImageCombo1.ComboItems(sParentFolder).Selected = True
                ListView1.ListItems.Clear
                Dir1.Path = sParentFolder
                ShowFolderList
                DeleteChild
            ElseIf (sParentFolder = "C:\WINDOWS\Desktop") Then
                ImageCombo1.ComboItems(1).Selected = True
                ListView1.ListItems.Clear
                Dir1.Path = sParentFolder
                ShowFolderList
                ShowFileList
                DeleteChild
            ElseIf (sParentFolder = "") Then
                'No Parent, Select My Computer
                ImageCombo1.ComboItems(2).Selected = True
                ListView1.ListItems.Clear
                Dir1.Path = sParentFolder
                GetMyComputer
            Else
                'Parent is a Folder
                ImageCombo1.ComboItems(sParentFolder).Selected = True
                ListView1.ListItems.Clear
                Dir1.Path = sParentFolder
                ShowFolderList
            End If
        Case "New"
            sKey = ImageCombo1.SelectedItem.Key
            AddNewFolder sKey, ""
            sKey = UCase(sKey)
            ImageCombo1_Click
            ListView1.SetFocus
            Set ListView1.SelectedItem = ListView1.ListItems(sKey & "New Folder")
            ListView1.SelectedItem.EnsureVisible
            ListView1.StartLabelEdit
        
    End Select
End Sub

Public Sub AddChild(sKeyChild As String, sTextChild As String)

Dim sParent As String

sParent = fso.GetParentFolderName(sKeyChild)
iNumOfChildren = iNumOfChildren + 1
ReDim Preserve Children(iNumOfChildren)
'Add Child
ChildIndex = iParentIndex + iNumOfChildren
If Len(sParent) > 3 Then 'Indent Subdirectories
    Set ci = ImageCombo1.ComboItems.Add(ChildIndex, sKeyChild, sTextChild, 2, 3, 4)
Else
    Set ci = ImageCombo1.ComboItems.Add(ChildIndex, sKeyChild, sTextChild, 2, 3, 3)
End If
    ci.Selected = True 'Select the Item in the ImageCombo
'Store Key in Array to use for delete
Children(iNumOfChildren) = sKeyChild


End Sub

Public Sub DeleteChild()

'Clear Previous children
For x = 1 To iNumOfChildren
sKey = Children(x)
ImageCombo1.ComboItems.Remove (sKey)
Children(x) = ""
Next

iNumOfChildren = 0
End Sub

Public Function CheckChild(Child As String) As Boolean
Dim sChildren As String

For x = 1 To iNumOfChildren
    sChildren = Children(x)
    If Child = sChildren Then
        CheckChild = True
        Exit Function
    Else
        CheckChild = False
    End If
Next

End Function

Private Sub CheckIcon(ByVal sFIle As String)

Dim sKey1 As String
Dim i As Long
Dim iHaveit As Long
Dim imgX As ListImage
Dim iPos As Long
Dim itmX As ListItem
Dim sExt As String
Dim lSicon As Object
Dim lLicon As Object
    ' We only want to get an icon for a given
    ' file type once, unless the file is an
    ' an executable or icon, in which case the
    ' icon is different for each instance of
    ' the extension type:

        sExt = UCase(fso.GetExtensionName(sFIle))
        If (sExt <> "EXE") And (sExt <> "ICO") And (sExt <> "DLL") And (sExt <> "OCX") And (sExt <> "HTML") And (sExt <> "LNK") And (sExt <> "") Then
            sKey1 = sExt
        Else
            sKey1 = sFIle
        End If
        sKey1 = UCase$(sKey1)
    ' Determine whether we've already got this type:
    For i = 1 To ilsIcons32.ListImages.Count
        If (ilsIcons32.ListImages(i).Key = sKey1) Then
            iHaveit = i
        End If
    Next i
    ' If we haven't already got it, then get the file
    ' icons and types and add them to the Image Lists:
    If (iHaveit = 0) Then
        
        Set imgX = ilsIcons32.ListImages.Add(, sKey1, GetIcon(sFIle, egitLargeIcon))
        imgX.Tag = GetFileTypeName(sFIle)
        iHaveit = imgX.Index
        ilsIcons16.ListImages.Add , sKey1, GetIcon(sFIle, egitSmallIcon)
        c = ilsIcons16.ListImages.Count
    End If
End Sub

Public Sub AddNewFolder(Path, FolderName)
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.GetFolder(Path)
  Set fc = f.SubFolders
  
  If FolderName <> "" Then
    Set nf = fc.Add(FolderName)
  Else
    Set nf = fc.Add("New Folder")
  End If
  Dir1.Refresh
End Sub

Public Sub ShowFolderList()
iListCount = ListView1.ListItems.Count
    For i = 0 To Dir1.ListCount - 1
        sKey = Dir1.List(i)
        sText = fso.GetBaseName(sKey)
        Set itmFldr = ListView1.ListItems.Add(iListCount + i + 1, sKey, sText, 1, 2)
    Next
    
End Sub

Public Sub ShowFileList()
On Error Resume Next
iListCount = ListView1.ListItems.Count
For i = 0 To File1.ListCount - 1
    If bIsDrive Then
        sKey = File1.Path & File1.List(i)
    Else
        sKey = File1.Path & "\" & File1.List(i)
    End If
    sText = File1.List(i)
    Length = Len(sText)
    'pos1 = InStr(Length - 5, sText, ".", vbBinaryCompare)
    sFType = fso.GetExtensionName(sKey) 'Mid(sText, pos1 + 1) ', 3)
    sFType = LCase(sFType)
    sExt = UCase(sFType)
    If sExt = "" Then
        sExt = "UNK"
    End If
    CheckIcon sKey
    If (sExt <> "EXE") And (sExt <> "ICO") And (sExt <> "DLL") And (sExt <> "OCX") And (sExt <> "HTML") And (sExt <> "LNK") And (sExt <> "UNK") Then
        sKey1 = sExt
    Else
        sKey1 = sKey
    End If
    sKey1 = UCase$(sKey1)
    ind1 = ilsIcons16.ListImages(sKey1).Index
    ind2 = ilsIcons32.ListImages(sKey1).Index
    Set itmFldr = ListView1.ListItems.Add(iListCount + i + 1, sKey, sText, sKey1, sKey1)
Next
End Sub

Public Sub GetlvReportData(folderspec)
On Error Resume Next
Dim sThis As String
Dim sThat As String
    If Len(folderspec) > 3 Then
        Set f = fso.GetFolder(folderspec)
        Set fc = f.Files
    Else
        Set f = fso.GetFolder(folderspec)
        Set fc = f.SubFolders
    End If
    
    For Each f1 In fc
        d = f1.DateCreated
        s = f1.Size
        t = f1.Type
        p = f1.Path
        sThis = Left$(p, 1)
        sThat = LCase$(sThis)
        p = Replace(p, sThis, sThat, 1, 1, vbBinaryCompare)
        Set itmX = ListView1.ListItems(p)
        itmX.SubItems(1) = s
        itmX.SubItems(2) = t
        itmX.SubItems(3) = d
    Next
    
End Sub

Public Sub GetDesktop()
On Error Resume Next
Screen.MousePointer = vbHourglass
ListView1.ListItems.Clear
Dir1.Path = "C:\WINDOWS\Desktop\"
File1.Path = Dir1.Path
iParentIndex = ImageCombo1.SelectedItem.Index
ShowFolderList
ShowFileList
Label2 = ListView1.ListItems.Count
Screen.MousePointer = vbDefault
End Sub

Public Sub GetMyComputer()
On Error Resume Next
iParentIndex = ImageCombo1.SelectedItem.Index
ListView1.ListItems.Clear
For i = 0 To Drive1.ListCount - 1
    d = Drive1.List(i)
    sText = d
    sKey = Left$(sText, 2) & "\"
    ind = ilsIcons16.ListImages(sKey).Index
    ListView1.ListItems.Add i + 1, sKey, sText, ind, ind
Next
End Sub
