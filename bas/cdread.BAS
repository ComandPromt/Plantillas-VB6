Attribute VB_Name = "CDCode"

Option Explicit

 

Declare Function mciSendString Lib "winmm" ()    As String    String, ByVal _uReturnLength As Long, ByVal hwndCallback As Long) _As Long







Type DiskTitleRecords

    IDNumber As Long

    Title As String * 128

    NumTracks As Integer

End Type



Global mciResult As Long

Global ReturnString As String * 128

Global ReturnString2 As String * 255

Global CDTitleFile As Integer

Global CurrentTitleRecordNumber As Integer

Global device As Boolean

Global sampling_rate As Long

Global bytes_per_second As Long

Global nchnls As Long

Global bits_per_sample As Long

Global align As Long

Global f$

Function mciMySendString(TheCommand As String) As Long

    

    Dim Dummy As Long

    Dummy = mciSendString(TheCommand, ReturnString, 127, 0)

    mciMySendString = Dummy

    

End Function





Function mciMySendString2(TheCommand As String) As Long

    

    Dim Dummy As Long

    Dummy = mciSendString(TheCommand, ReturnString2, 256, 0)

    mciMySendString2 = Dummy

    

End Function

Function CleanString(MessyString As String) As String



    CleanString = Left$(MessyString, InStr(MessyString, Chr$(0)) - 1)



End Function



Function GetDiskTitleRecord(AnyDiskTitleRecord As DiskTitleRecords)



    Dim NumberOfTracks As Integer

    Dim DiskLength As Long

    Dim DiskID As Long

    Dim RecordPosition As Integer

    Dim Found As Integer

    

    mciResult = mciMySendString("status CDAudio number of tracks")

    NumberOfTracks = Val(CleanString(ReturnString))

    mciResult = mciMySendString("set CDAudio time format milliseconds")

    mciResult = mciMySendString("status CDAudio length")

    DiskLength = Val(CleanString(ReturnString))

    DiskID = NumberOfTracks * DiskLength

    

    RecordPosition = 0

    

    Do

        RecordPosition = RecordPosition + 1

        Get CDTitleFile, RecordPosition, AnyDiskTitleRecord

    Found = (DiskID = AnyDiskTitleRecord.IDNumber)

    Loop Until Found Or EOF(CDTitleFile)

    

    If Found Then

        GetDiskTitleRecord = RecordPosition

    Else

        GetDiskTitleRecord = 0

        AnyDiskTitleRecord.IDNumber = DiskID

        AnyDiskTitleRecord.NumTracks = NumberOfTracks

    End If

    

End Function



Sub OpenTitleFile()



    Dim AnyTitleRecord As DiskTitleRecords

    

    CDTitleFile = 1

    

    Open App.Path & "\CDTitles.Dat" For Random As CDTitleFile _

        Len = Len(AnyTitleRecord)

    

    Open "a:\CDTitles.Dat" For Random As CDTitleFile _

        Len = Len(AnyTitleRecord)

      

End Sub



Sub SaveDiskTitleRecord(AnyDiskTitleRecord As DiskTitleRecords, _

    RecordPosition As Integer)

    

    Dim FileSize As Integer

    

    If RecordPosition <> 0 Then

        Put CDTitleFile, RecordPosition, AnyDiskTitleRecord

    Else

        FileSize = LOF(CDTitleFile) \ Len(AnyDiskTitleRecord)

        RecordPosition = FileSize + 1

        Put CDTitleFile, RecordPosition, AnyDiskTitleRecord

    End If

    

End Sub



