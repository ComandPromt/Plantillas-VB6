Attribute VB_Name = "modMake"
'If you are going to use this in a app, you must
'first contact me at aandrei@hades.ro, and you
'have to credit me on the application's box, and/or
'about box

Sub AddToSelfExtract(SelfExtract As String, WhatFile As String, SaveAs As String)

Dim iFreeFile As Integer
Dim iFreeFile2 As Integer
Dim sBuffer As String
Dim sBefore As String

iFreeFile = FreeFile

Open SelfExtract For Binary As iFreeFile
    sBefore = String(LOF(iFreeFile), Chr(0))
    Get iFreeFile, , sBefore
Close iFreeFile
    

Open SaveAs For Output As iFreeFile
    iFreeFile2 = FreeFile
    Open WhatFile For Binary As iFreeFile2
        sBuffer = String(LOF(iFreeFile2), Chr(0))
        Get iFreeFile2, , sBuffer
        Size = LOF(iFreeFile2)
        Size = String(10 - Len(Size), "0") & Size
        Print #iFreeFile, sBefore & sBuffer & Size
    Close iFreeFile2
Close iFreeFile

End Sub
