Imports System.IO
Imports System.Text
Imports System.Math
Imports System.Threading

Public Class MP3Form
    Inherits System.Windows.Forms.Form
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents picCover As System.Windows.Forms.PictureBox
    Friend WithEvents btnFList As System.Windows.Forms.Button
    Friend WithEvents txtStartDir As System.Windows.Forms.TextBox
    Friend WithEvents lstTrackInfo As System.Windows.Forms.ListBox
    Friend WithEvents lstPlayList As System.Windows.Forms.ListBox
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnDel As System.Windows.Forms.Button
    Friend WithEvents btnUp As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dlgSave As System.Windows.Forms.SaveFileDialog
    Friend WithEvents btnWalkDir As System.Windows.Forms.Button
    Friend WithEvents dlgDir As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents lvwFNames As System.Windows.Forms.ListView
    Friend WithEvents colTitle As System.Windows.Forms.ColumnHeader
    Friend WithEvents colAlbum As System.Windows.Forms.ColumnHeader
    Friend WithEvents colArtist As System.Windows.Forms.ColumnHeader
    Friend WithEvents colFName As System.Windows.Forms.ColumnHeader
    Friend WithEvents txtStatus As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.picCover = New System.Windows.Forms.PictureBox
        Me.btnFList = New System.Windows.Forms.Button
        Me.txtStartDir = New System.Windows.Forms.TextBox
        Me.lstTrackInfo = New System.Windows.Forms.ListBox
        Me.lstPlayList = New System.Windows.Forms.ListBox
        Me.btnAdd = New System.Windows.Forms.Button
        Me.btnDel = New System.Windows.Forms.Button
        Me.btnUp = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.dlgSave = New System.Windows.Forms.SaveFileDialog
        Me.btnWalkDir = New System.Windows.Forms.Button
        Me.dlgDir = New System.Windows.Forms.FolderBrowserDialog
        Me.lvwFNames = New System.Windows.Forms.ListView
        Me.colTitle = New System.Windows.Forms.ColumnHeader
        Me.colArtist = New System.Windows.Forms.ColumnHeader
        Me.colAlbum = New System.Windows.Forms.ColumnHeader
        Me.colFName = New System.Windows.Forms.ColumnHeader
        Me.txtStatus = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'picCover
        '
        Me.picCover.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.picCover.Location = New System.Drawing.Point(568, 160)
        Me.picCover.Name = "picCover"
        Me.picCover.Size = New System.Drawing.Size(208, 192)
        Me.picCover.TabIndex = 3
        Me.picCover.TabStop = False
        '
        'btnFList
        '
        Me.btnFList.BackColor = System.Drawing.Color.Silver
        Me.btnFList.Location = New System.Drawing.Point(8, 16)
        Me.btnFList.Name = "btnFList"
        Me.btnFList.Size = New System.Drawing.Size(80, 24)
        Me.btnFList.TabIndex = 4
        Me.btnFList.Text = "Walk Files"
        '
        'txtStartDir
        '
        Me.txtStartDir.Location = New System.Drawing.Point(96, 16)
        Me.txtStartDir.Name = "txtStartDir"
        Me.txtStartDir.Size = New System.Drawing.Size(192, 20)
        Me.txtStartDir.TabIndex = 6
        Me.txtStartDir.Text = "C:\music"
        '
        'lstTrackInfo
        '
        Me.lstTrackInfo.Location = New System.Drawing.Point(544, 48)
        Me.lstTrackInfo.Name = "lstTrackInfo"
        Me.lstTrackInfo.Size = New System.Drawing.Size(256, 95)
        Me.lstTrackInfo.TabIndex = 7
        '
        'lstPlayList
        '
        Me.lstPlayList.AllowDrop = True
        Me.lstPlayList.ColumnWidth = 400
        Me.lstPlayList.Location = New System.Drawing.Point(8, 392)
        Me.lstPlayList.MultiColumn = True
        Me.lstPlayList.Name = "lstPlayList"
        Me.lstPlayList.Size = New System.Drawing.Size(728, 212)
        Me.lstPlayList.TabIndex = 8
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.Color.Silver
        Me.btnAdd.Location = New System.Drawing.Point(8, 360)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(48, 24)
        Me.btnAdd.TabIndex = 9
        Me.btnAdd.Text = "Add"
        '
        'btnDel
        '
        Me.btnDel.BackColor = System.Drawing.Color.Silver
        Me.btnDel.Location = New System.Drawing.Point(752, 456)
        Me.btnDel.Name = "btnDel"
        Me.btnDel.Size = New System.Drawing.Size(48, 24)
        Me.btnDel.TabIndex = 10
        Me.btnDel.Text = "Del"
        '
        'btnUp
        '
        Me.btnUp.BackColor = System.Drawing.Color.Silver
        Me.btnUp.Location = New System.Drawing.Point(752, 392)
        Me.btnUp.Name = "btnUp"
        Me.btnUp.Size = New System.Drawing.Size(48, 24)
        Me.btnUp.TabIndex = 11
        Me.btnUp.Text = "Up"
        '
        'btnDown
        '
        Me.btnDown.BackColor = System.Drawing.Color.Silver
        Me.btnDown.Location = New System.Drawing.Point(752, 424)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(48, 24)
        Me.btnDown.TabIndex = 12
        Me.btnDown.Text = "Down"
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.Color.Silver
        Me.btnSave.Location = New System.Drawing.Point(752, 512)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(48, 24)
        Me.btnSave.TabIndex = 13
        Me.btnSave.Text = "Save..."
        '
        'dlgSave
        '
        Me.dlgSave.Filter = "Playlist (*.m3u)|*.m3u|All files (*.*)|*.*"
        '
        'btnWalkDir
        '
        Me.btnWalkDir.BackColor = System.Drawing.Color.Silver
        Me.btnWalkDir.Location = New System.Drawing.Point(296, 16)
        Me.btnWalkDir.Name = "btnWalkDir"
        Me.btnWalkDir.Size = New System.Drawing.Size(24, 24)
        Me.btnWalkDir.TabIndex = 14
        Me.btnWalkDir.Text = "..."
        '
        'lvwFNames
        '
        Me.lvwFNames.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colTitle, Me.colArtist, Me.colAlbum, Me.colFName})
        Me.lvwFNames.Location = New System.Drawing.Point(8, 48)
        Me.lvwFNames.Name = "lvwFNames"
        Me.lvwFNames.Size = New System.Drawing.Size(528, 304)
        Me.lvwFNames.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lvwFNames.TabIndex = 15
        Me.lvwFNames.View = System.Windows.Forms.View.Details
        '
        'colTitle
        '
        Me.colTitle.Text = "Song"
        Me.colTitle.Width = 153
        '
        'colArtist
        '
        Me.colArtist.Text = "Artist"
        Me.colArtist.Width = 143
        '
        'colAlbum
        '
        Me.colAlbum.Text = "Album"
        Me.colAlbum.Width = 147
        '
        'colFName
        '
        Me.colFName.Text = "File"
        Me.colFName.Width = 161
        '
        'txtStatus
        '
        Me.txtStatus.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.txtStatus.Location = New System.Drawing.Point(344, 16)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.Size = New System.Drawing.Size(240, 16)
        Me.txtStatus.TabIndex = 16
        '
        'MP3Form
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.RoyalBlue
        Me.ClientSize = New System.Drawing.Size(808, 630)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.lvwFNames)
        Me.Controls.Add(Me.btnWalkDir)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnDown)
        Me.Controls.Add(Me.btnUp)
        Me.Controls.Add(Me.btnDel)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.lstPlayList)
        Me.Controls.Add(Me.lstTrackInfo)
        Me.Controls.Add(Me.txtStartDir)
        Me.Controls.Add(Me.btnFList)
        Me.Controls.Add(Me.picCover)
        Me.Name = "MP3Form"
        Me.Text = "MSDN Magazine MP3 Reader"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim colMP3 As Collection

    Private Sub btnFList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFList.Click
        Dim S As MP3

        '        MP3.WalkDirs(txtStartDir.Text, colMP3)
        '        MP3.WalkMP3FileNames(txtStartDir.Text, colMP3)
        lvwFNames.Items.Clear()

        ' WalkMP3FileNames recurses from the given directory, finding all *.mp3 files and
        ' adding their filenames to the appropriate ListView.

        MP3.WalkMP3FileNames(txtStartDir.Text, lvwFNames)

        txtStatus.Text = "Read " & lvwFNames.Items.Count & " songs. Filling in..."
        Dim T As New Thread(AddressOf FillInListViewItems)
        T.Start()

    End Sub

    Sub FillInListViewItems()
        Dim lvi As ListViewItem
        Dim n As Integer

        n = 0
        For Each lvi In lvwFNames.Items
            Dim M As New MP3(lvi.SubItems(3).Text)
            lvi.SubItems(0).Text = M.Title
            lvi.SubItems(1).Text = M.Artist
            lvi.SubItems(2).Text = M.Album
            n += 1
            txtStatus.Text = "Filled in " & n & " of " & lvwFNames.Items.Count
            lvwFNames.Refresh()
            M = Nothing
        Next
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        lstPlayList.Items.Add(lvwFNames.SelectedItems(0).SubItems(3).Text)
    End Sub

    Private Sub btnUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUp.Click
        ' Swap the current item with the one previous to it.
        Dim lbprv As Object

        If lstPlayList.SelectedIndex = 0 Then Return

        lbprv = lstPlayList.Items(lstPlayList.SelectedIndex - 1)
        lstPlayList.Items(lstPlayList.SelectedIndex - 1) = lstPlayList.Items(lstPlayList.SelectedIndex)
        lstPlayList.Items(lstPlayList.SelectedIndex) = lbprv
        lstPlayList.SelectedIndex -= 1
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        ' Swap the current item with the one previous to it.
        Dim lbprv As Object

        If lstPlayList.SelectedIndex = lstPlayList.Items.Count - 1 Then Return

        lbprv = lstPlayList.Items(lstPlayList.SelectedIndex + 1)
        lstPlayList.Items(lstPlayList.SelectedIndex + 1) = lstPlayList.Items(lstPlayList.SelectedIndex)
        lstPlayList.Items(lstPlayList.SelectedIndex) = lbprv
        lstPlayList.SelectedIndex += 1

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim FSave As String
        Dim hFSave As StreamWriter
        Dim SongPath As String

        dlgSave.ShowDialog()
        FSave = dlgSave.FileName

        ' Open MP3 File
        hFSave = New StreamWriter(FSave, False) ' False - don't append, recreate.

        For Each SongPath In lstPlayList.Items
            hFSave.WriteLine(SongPath.Substring(2))  ' Start at 3d char - trim drive letter off!
        Next

        hFSave.Close()
    End Sub

    Private Sub btnWalkDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWalkDir.Click
        dlgDir.ShowDialog()
        txtStartDir.Text = dlgDir.SelectedPath
    End Sub

    Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
        lstPlayList.Items.Remove(lstPlayList.SelectedItem)
    End Sub

    Private Sub lvwFNames_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvwFNames.SelectedIndexChanged
        Dim lvwSong As ListViewItem
        Try
            lvwSong = lvwFNames.FocusedItem()  'SelectedItems(0)
            Dim S As New MP3(lvwSong.SubItems(3).Text)

            AddSongInfoToListBox(S)
        Catch ex As Exception
        End Try
    End Sub

    ' =============================================================
    ' AddSongInfoToListBox
    ' 

    Sub AddSongInfoToListBox(ByVal theSong As MP3)
        lstTrackInfo.Items.Clear()
        lstTrackInfo.Items.Add("Track length (ms): " & msToTimeStr(theSong.msec))
        lstTrackInfo.Items.Add("Title: " & theSong.Title)
        lstTrackInfo.Items.Add("Track: " & theSong.Track)
        lstTrackInfo.Items.Add("Artist: " & theSong.Artist)
        lstTrackInfo.Items.Add("Genre: " & theSong.Genre)
        lstTrackInfo.Items.Add("Album:  " & theSong.Album)

        Try
            picCover.Image = Nothing
            picCover.Image = theSong.GetCoverImage()
            picCover.Width = picCover.Image.Width
            picCover.Height = picCover.Image.Height
        Catch ex As Exception
        End Try
    End Sub

    ' =============================================================
    ' msToTimeStr
    '
    ' 

    Function msToTimeStr(ByVal ms As Long) As String
        Dim TimeString As String
        Dim Sec As Single, Minu As Single
        Dim SStr As String

        Minu = (ms / 60000)
        Sec = (Minu - Floor(Minu)) * 60
        TimeString = Floor(Minu) & ":" & CStr(Round(Sec)).PadLeft(2, "0")

        Return TimeString
    End Function

End Class
