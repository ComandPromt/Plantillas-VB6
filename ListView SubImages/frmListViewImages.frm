VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListViewImages 
   Caption         =   "ListView With Images"
   ClientHeight    =   3330
   ClientLeft      =   615
   ClientTop       =   1395
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   11505
   Begin MSComctlLib.ImageList imglstListImages 
      Left            =   10680
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListViewImages.frx":0000
            Key             =   "S"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListViewImages.frx":015C
            Key             =   "Flag"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListViewImages.frx":02B8
            Key             =   "Clip"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListViewImages.frx":0414
            Key             =   "A"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListViewImages.frx":0570
            Key             =   "Bolt"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListViewImages.frx":06CC
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListViewImages.frx":0828
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListViewImages.frx":0984
            Key             =   "Mail_New"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListViewImages.frx":0AE0
            Key             =   "Mail_Read"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListViewImages.frx":0C3C
            Key             =   "R"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstListView 
      Height          =   3315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   5847
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imglstListImages"
      SmallIcons      =   "imglstListImages"
      ColHdrIcons     =   "imglstListImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmListViewImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    ' build our column headers
    With lstListView.ColumnHeaders
        
        .Add , , "", 400, , "S"
        .Add , , "", 400, , "Flag"
        .Add , , "", 400, , "Clip"
        .Add , , "", 400, , "A"
        .Add , , "", 400, , "Bolt"
        .Add , , "Subject", 3000
        .Add , , "Date", 1000, , "Down"
        .Add , , "Time", 1000
        .Add , , "Size", 1000
        
    End With
    
    ' set the default sort to the date column
    lstListView.Sorted = True
    lstListView.SortOrder = lvwAscending
    lstListView.SortKey = 5
    
    ' just to show this off
    LoadList
    
End Sub

Private Sub Form_Resize()
    
    ' ignore resize errors
    On Error Resume Next
    
    ' resize the listview
    lstListView.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
End Sub

Private Sub lstListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    ' make sure the list has sorting enabled
    lstListView.Sorted = True
    
    If lstListView.SortOrder = lvwAscending Then
        ' if it's currently asc then desc it
        lstListView.SortOrder = lvwDescending
    Else
        ' if it's currently desc then asc it
        lstListView.SortOrder = lvwAscending
    End If
    
    ' only change the icon if there isn't one there
    ' or it's an arrow
    If lstListView.ColumnHeaders(ColumnHeader.Index).Icon = 0 Or _
    lstListView.ColumnHeaders(ColumnHeader.Index).Icon = "Up" Then
        lstListView.ColumnHeaders(ColumnHeader.Index).Icon = "Down"
        GoTo ClearAllOthers
    End If
    
    ' only change the icon if there isn't one there
    ' or it's an arrow
    If lstListView.ColumnHeaders(ColumnHeader.Index).Icon = 0 Or _
    lstListView.ColumnHeaders(ColumnHeader.Index).Icon = "Down" Then
        lstListView.ColumnHeaders(ColumnHeader.Index).Icon = "Up"
        GoTo ClearAllOthers
    End If
    
ClearAllOthers:
    ' setup a counter variable
    Dim lngIndex As Long
    
    ' loop through all of the column headers
    For lngIndex = 1 To lstListView.ColumnHeaders.Count - 1
        ' except the current one
        If lngIndex <> ColumnHeader.Index Then
            ' and if it has an 'up' or 'down' then
            If lstListView.ColumnHeaders(lngIndex).Icon = "Up" Or _
            lstListView.ColumnHeaders(lngIndex).Icon = "Down" Then
                ' dectroy it's icon
                lstListView.ColumnHeaders(lngIndex).Icon = 0
            End If
        End If
    Next lngIndex
    
End Sub

Private Sub LoadList()
    
    ' add some junk items to the list
    ' see function 'AddToListView' for the details
    AddToListView "Re: Feliz Navidad", "22/12/99", "0:18", "8K", False, True
    AddToListView "Feliz Navidad", "20/12/99", "13:02", "263K", True, False, True
    AddToListView "Feliz Navid...aggggg", "13/12/99", "22:02", "28K", True
    AddToListView "Re; web", "19/11/99", "18:38", "5K", True
    AddToListView "Rases de datos de liber-Swiss", "05/05/99", "9:59", "865K", True, True
    AddToListView "Re; Boomerang... Contestacion", "02/04/99", "11:16", "5K", True
    AddToListView "RV:", "12/03/99", "2:41", "676K", True, True
    AddToListView "RV: Esto es BUENISIMO...!!!", "27/02/99", "9:24", "3K", True
    
    ' refresh the lsitview and select the first item
    lstListView.Refresh
    lstListView.ListItems.Item(1).Selected = True
    
End Sub

Private Function AddToListView(strSubject As String, _
                               strDate As String, _
                               strTime As String, _
                               strSize As String, _
                               Optional bolMailRead As Boolean, _
                               Optional bolPaperClip As Boolean, _
                               Optional bolShowR As Boolean)
    
    ' Note: i did not handle the second (flag) and fifth (bolt) columns
    ' because the image you sent sisn't have them in it. But I'm sure
    ' that from this code you can figure it out easily.
    
    ' setup a variable to use to build an item
    Dim lstEntry As ListItem
    
    ' set the icon to the first item
    If bolMailRead = True Then
        Set lstEntry = lstListView.ListItems.Add(, , " ", "Mail_Read")
    Else
        Set lstEntry = lstListView.ListItems.Add(, , " ", "Mail_New")
    End If
    
    ' build the first two children
    lstEntry.SubItems(1) = ""
    lstEntry.SubItems(2) = ""
    
    ' if we wanted a 'clip' then set the clip
    If bolPaperClip = True Then lstEntry.ListSubItems.Item(2).ReportIcon = "Clip"
    
    ' build another blank child
    lstEntry.SubItems(3) = ""
    
    ' if we wanted an 'r' then set the r
    If bolShowR = True Then lstEntry.ListSubItems.Item(3).ReportIcon = "R"
    
    ' build another blank child
    lstEntry.SubItems(4) = ""
    
    ' set the items to teh desired values
    lstEntry.SubItems(5) = strSubject
    lstEntry.SubItems(6) = strDate
    lstEntry.SubItems(7) = strTime
    lstEntry.SubItems(8) = strSize
    
End Function
