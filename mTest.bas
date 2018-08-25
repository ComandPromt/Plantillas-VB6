Attribute VB_Name = "mTest"
Option Explicit
'****************************************************************
'*  VB file:   mTest.bas
'*
'*  DEMO OF CALLING FUNCTIONS IN GetFree.bas
'*
'*  Copyright (c) 1998, Ray Mercer.  All rights reserved.
'****************************************************************
Sub Main()
Dim sMsg As String
Dim f As Boolean

sMsg = "INFO ON CURRENT DRIVE:" & vbCrLf
sMsg = sMsg & "Free bytes: " & vbGetAvailableBytesAsString() & vbCrLf
sMsg = sMsg & "Free kilobytes: " & vbGetAvailableKBytesAsString() & vbCrLf
sMsg = sMsg & "Free megabytes: " & vbGetAvailableMBytesAsString() & vbCrLf
sMsg = sMsg & "Total bytes: " & vbGetTotalBytesAsString() & vbCrLf
sMsg = sMsg & "Total kilobytes: " & vbGetTotalKBytesAsString() & vbCrLf
sMsg = sMsg & "Total megabytes: " & vbGetTotalMBytesAsString() & vbCrLf
sMsg = sMsg & vbCrLf
sMsg = sMsg & "Percentage of DiskSpaceFree: " & vbGetPercentAvailable() & "%" & vbCrLf
sMsg = sMsg & vbCrLf

f = ExistGetDiskFreeSpaceEx()
sMsg = sMsg & "GetDiskFreeSpaceEx() Available = " & f & vbCrLf

Dim userRet As VbMsgBoxResult

userRet = MsgBox(sMsg, vbOKOnly, App.Title)

End Sub
