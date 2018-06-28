Imports System.IO
Imports System.Text

Public Class MP3

    Public Path As String
    Public Title As String
    Public Album As String
    Public Artist As String
    Public Genre As String
    Public msec As Long
    Public Track As Integer

    Public Structure Tag
        Dim MTag As String
        Dim MFlag As Integer
        Dim MSize As Integer
        Dim MData As Byte()
        Dim StrData As String
    End Structure

    ' =============================================================
    ' MP3.WalkDirs
    '
    ' Given a starting directory, walks the tree and adds all MP3 
    ' files to a string collection.
    '

    Public Shared Function WalkDirs(ByVal StartDir As String, ByVal colMP3 As Collection)
        Dim DI As Directory
        Dim AllDirs As String()
        Dim D As String
        Dim M As String
        Dim AllMP3s As String()
        Dim S As MP3

        AllDirs = DI.GetDirectories(StartDir)

        Try
            For Each D In AllDirs
                WalkDirs(D, colMP3)
            Next
        Catch
        End Try

        AllMP3s = DI.GetFiles(StartDir, "*.mp3")

        For Each M In AllMP3s

            ' Get song/artist/album/file
            Dim MTags As Collection

            S = New MP3
            S.Path = M

            MTags = S.TagsFromFile() ' This needs just the Path value

            Try
                S.Title = MTags("TIT2").StrData
                S.Album = MTags("TALB").StrData
                S.Artist = MTags("TPE1").StrData
                S.msec = CLng(MTags("TLEN").StrData)
                S.Genre = MTags("TCON").StrData
                S.Track = CInt(MTags("TRCK").StrData)
            Catch ex As Exception

            End Try

            colMP3.Add(S, S.Path)

            Application.DoEvents()

        Next

    End Function

    Public Shared Function WalkMP3FileNames(ByVal StartDir As String, Optional ByVal Frm As ListView = Nothing)
        Dim DI As Directory
        Dim AllDirs As String()
        Dim D As String
        Dim M As String
        Dim AllMP3s As String()
        Dim S As MP3

        AllDirs = DI.GetDirectories(StartDir)

        Try
            For Each D In AllDirs
                WalkMP3FileNames(D, Frm)
            Next
        Catch
        End Try

        AllMP3s = DI.GetFiles(StartDir, "*.mp3")

        For Each M In AllMP3s

            If Not IsNothing(Frm) Then
                Dim lvwSong As ListViewItem

                lvwSong = New ListViewItem("")
                lvwSong.SubItems.Add("")
                lvwSong.SubItems.Add("")
                lvwSong.SubItems.Add(M)

                Frm.Items.Add(lvwSong)
                Frm.Refresh()

            End If
        Next

    End Function

    ' =============================================================
    ' TagsFromFile
    '
    ' Reads the MP3 tags from a file, given the file's pathname.
    ' Returns the tags as a collection.
    '

    Private Function TagsFromFile() As Collection
        Dim sMPEGFile As FileStream
        Dim ba(15) As Byte
        Dim theTags As New Collection

        Try
            ' Open MP3 File
            sMPEGFile = New FileStream(Path, FileMode.Open)

            ' Read MP3 header info
            sMPEGFile.Read(ba, 0, 10)

            If (ba(0) = &H49) And (ba(1) = &H44) And (ba(2) = &H33) Then ' It's a proper MP3

            End If

            Dim theTag As Tag

            theTag = TagFromStream(sMPEGFile)

            Do While Not IsNothing(theTag.MTag)
                theTags.Add(theTag, theTag.MTag)
                theTag = New Tag
                theTag = TagFromStream(sMPEGFile)
            Loop

        Catch ex As Exception
        End Try

        sMPEGFile.Close()
        Return theTags
    End Function


    ' =============================================================
    ' TagFromStream
    '
    ' Reads the next tag from an MP3 file, represented by a stream
    '

    Private Shared Function TagFromStream(ByVal sMPEG As Stream) As Tag
        Dim theTag As New Tag

        Dim ba() As Byte
        Dim ascii As New ASCIIEncoding

        ReDim ba(5)
        Try
            sMPEG.Read(ba, 0, 4)
            theTag.MTag = ascii.GetString(ba)

            If ba(0) = 0 Then Return Nothing

            sMPEG.Read(ba, 0, 4)
            theTag.MSize = (65536 * (ba(0) * 256 + ba(1))) + (ba(2) * 256 + ba(3))

            sMPEG.Read(ba, 0, 3)
            theTag.MFlag = ba(0) * 256 + ba(1)

            ReDim ba(theTag.MSize + 1)

            If theTag.MSize > 0 Then
                sMPEG.Read(ba, 0, theTag.MSize - 1)
                theTag.MData = ba

                If theTag.MTag.Substring(0, 1) = "T" Then
                    theTag.StrData = ascii.GetString(ba)
                End If
            End If
        Catch ex As Exception
        End Try

        Return theTag
    End Function

    ' =============================================================
    ' GetCoverImage
    '
    ' Given a Song, opens the path name and retrieves the cover 
    ' image if it exists.
    '

    Public Function GetCoverImage() As Image
        Dim APic As Byte()
        Dim theTags

        Try
            theTags = Me.TagsFromFile()
            APic = theTags("APIC").MData
        Catch ex As Exception
        End Try

        Return MP3.GetImageFromByteArray(APic)

    End Function

    ' =============================================================
    ' GetImageFromByteArray
    '
    ' Converts a Byte() into an Image
    '

    Private Shared Function GetImageFromByteArray(ByVal bdata As Byte()) As Image
        Dim zidx As Integer
        Dim jpgdata As Byte()
        Dim n As Integer

        zidx = bdata.IndexOf(bdata, 0, 1)
        zidx += 2

        zidx = 12
        ReDim jpgdata(bdata.Length)

        Array.Copy(bdata, 12, jpgdata, 0, bdata.Length - 13)

        Return Image.FromStream(New MemoryStream(jpgdata))

    End Function

    Public Sub New()

    End Sub

    Public Sub New(ByVal Path As String)

        Me.Path = Path

        Dim MTags As Collection

        Try
            MTags = TagsFromFile() ' This needs just the Path value
            Me.Title = MTags("TIT2").StrData
            Me.Album = MTags("TALB").StrData
            Me.Artist = MTags("TPE1").StrData
            Me.msec = CLng(MTags("TLEN").StrData)
            Me.Genre = MTags("TCON").StrData
            Me.Track = CInt(MTags("TRCK").StrData)
        Catch ex As Exception

        End Try

    End Sub

    ' =============================================================
    ' MakeTagList
    '
    ' Creates a collection of tag names, based on official ID3 designations.
    ' This can be used to look up a tag that's embedded in an MP3 file.
    '
    ' Info from http://www.id3.org/id3v2.3.0.txt
    '

    Public Shared Function MakeTagList() As Collection

        Dim TagList As New Collection

        TagList.Add("Audio Encryption", "AENC")
        TagList.Add("Attached picture", "APIC")
        TagList.Add("Comments", "COMM")
        TagList.Add("Commercial frame", "COMR")
        TagList.Add("Encryption method registration", "ENCR")
        TagList.Add("Equalization", "EQUA")
        TagList.Add("Event timing codes", "ETCO")
        TagList.Add("General encapsulated object", "GEOB")
        TagList.Add("Group identification registration", "GRID")
        TagList.Add("Involved people list", "IPLS")
        TagList.Add("Linked information", "LINK")
        TagList.Add("Music CD identifier", "MCDI")
        TagList.Add("MPEG location lookup table", "MLLT")
        TagList.Add("Ownership frame", "OWNE")
        TagList.Add("Private frame", "PRIV")
        TagList.Add("Play counter", "PCNT")
        TagList.Add("Popularimeter", "POPM")
        TagList.Add("Position synchronisation frame", "POSS")
        TagList.Add("Recommended buffer size", "RBUF")
        TagList.Add("Relative volume adjustment", "RVAD")
        TagList.Add("Reverb", "RVRB")
        TagList.Add("Synchronized lyric/text", "SYLT")
        TagList.Add("Synchronized tempo codes", "SYTC")
        TagList.Add("Album/Movie/Show title", "TALB")
        TagList.Add("BPM (beats per minute)", "TBPM")
        TagList.Add("Composer", "TCOM")
        TagList.Add("Content type", "TCON")
        TagList.Add("Copyright message", "TCOP")
        TagList.Add("Date", "TDAT")
        TagList.Add("Playlist delay", "TDLY")
        TagList.Add("Encoded by", "TENC")
        TagList.Add("Lyricist/Text writer", "TEXT")
        TagList.Add("File type", "TFLT")
        TagList.Add("Time", "TIME")
        TagList.Add("Content group description", "TIT1")
        TagList.Add("Title/songname/content description", "TIT2")
        TagList.Add("Subtitle/Description refinement", "TIT3")
        TagList.Add("Initial key", "TKEY")
        TagList.Add("Language(s)", "TLAN")
        TagList.Add("Length", "TLEN")
        TagList.Add("Media type", "TMED")
        TagList.Add("Original album/movie/show title", "TOAL")
        TagList.Add("Original filename", "TOFN")
        TagList.Add("Original lyricist(s)/text writer(s)", "TOLY")
        TagList.Add("Original artist(s)/performer(s)", "TOPE")
        TagList.Add("Original release year", "TORY")
        TagList.Add("File owner/licensee", "TOWN")
        TagList.Add("Lead performer(s)/Soloist(s)", "TPE1")
        TagList.Add("Band/orchestra/accompaniment", "TPE2")
        TagList.Add("Conductor/performer refinement", "TPE3")
        TagList.Add("Interpreted, remixed, or otherwise modified by", "TPE4")
        TagList.Add("Part of a set", "TPOS")
        TagList.Add("Publisher", "TPUB")
        TagList.Add("Track number/Position in set", "TRCK")
        TagList.Add("Recording dates", "TRDA")
        TagList.Add("Internet radio station name", "TRSN")
        TagList.Add("Internet radio station owner", "TRSO")
        TagList.Add("Size", "TSIZ")
        TagList.Add("ISRC (international standard recording code)", "TSRC")
        TagList.Add("Software/Hardware and settings used for encoding", "TSSE")
        TagList.Add("Year", "TYER")
        TagList.Add("User defined text information frame", "TXXX")
        TagList.Add("Unique file identifier", "UFID")
        TagList.Add("Terms of use", "USER")
        TagList.Add("Unsychronized lyric/text transcription", "USLT")
        TagList.Add("Commercial information", "WCOM")
        TagList.Add("Copyright/Legal information", "WCOP")
        TagList.Add("Official audio file webpage", "WOAF")
        TagList.Add("Official artist/performer webpage", "WOAR")
        TagList.Add("Official audio source webpage", "WOAS")
        TagList.Add("Official internet radio station homepage", "WORS")
        TagList.Add("Payment", "WPAY")
        TagList.Add("Publishers official webpage", "WPUB")
        TagList.Add("User defined URL link frame", "WXXX")

        Return TagList
    End Function

End Class
