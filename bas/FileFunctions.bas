Attribute VB_Name = "FileFunctions"
' File Function Library
' Copyright 1996 Jens Balchen
'
' Module dependencies
'

Const MODULE_NAME = "File function library"

Option Explicit

Function FileExists%(filename$)

' Description
'     Checks 'filename' to find wether the filename given
'     exists.
'
' Parameters
'     Name              Type     Value
'     -------------------------------------------------------------
'     filename$         String   The filename to be checked
'
' Returns
'     True if the file exists
'     False if the file does not exist
'
' Last updated by Jens Balchen 1996-03-16

Dim f%

   ' Trap any errors that may occur
   On Error Resume Next

   ' Get a free file handle to avoid using a file handle already in use
   f% = FreeFile
   ' Open the file for reading
   Open filename$ For Input As #f%
   ' Close it
   Close #f%
   ' If there was an error, Err will be <> 0. In that case, we return False
   FileExists% = Not (Err <> 0)

End Function

Sub KillFile(ByVal filename$)

' Description
'     Kills a file. If the file doesn't exist, the error
'     is trapped and disgarded
'
' Parameters
'     Name                    Type        Value
'     ----------------------------------------------------
'     filename                String      The file to kill
'
' Returns
'     Nothing
'
' Last modified by Jens Balchen 1996-06-30

   ' Trap any errors
   On Error Resume Next
   
   ' Delete file
   Kill filename$
   
   ' Reset error trapping to no trapping
   On Error GoTo 0

End Sub

