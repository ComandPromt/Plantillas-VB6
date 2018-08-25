Attribute VB_Name = "IO"
Option Explicit
    ' Procedure:
    '   CreateDir - Function
    '
    '   this function calls itself
    '   recurse through directory path
    '   creating new directorys that are
    '   separated by the '\' character
    '
    ' author:
    '       Richard G.Dormer
    '
    ' arguments:
    '       path - path to create
    '
    ' use:
    ' dim returnval as boolean
    '   1) returnval = CreateDir( "c:\apath\to\success" )
    '   2) if CreateDir( "c:\apath\to\success" ) then
    '         ' do something
    '      else
    '         ' do something else
    '      endif
    '
    ' variables:
    '   start - where the InStr starts searching for "\"
    '   pos - the return value of InStr
    '   directory - the current directory (buffer)
    '   result - the ultimate results of the function
    '
    ' return:
    '   true if successful; else not true
    '
    '
    ' note you can cut and paste this function
    ' out of the IO.bas file and into yours
    ' If you choose to leave it. You must use
    ' the module Identifier before the function name
    ' I.E  returnval = IO.CreateDir( "c:\apath\to\success" )
    '
Public Function CreateDir(path As String) As Boolean
 
    Static start, pos As Integer
    Static directory As String
    Static result As Boolean
    result = True
    
    ' initialize the error trap
    On Error GoTo errCreation
    
    ' if null string why bother....
    If path = "" Then Err.Raise vbObjectError + 1
    
    ' start will always be null
    ' the first time through
    If start = Empty Then
        start = 1
    Else
         start = pos + 1
    End If
                
    ' find "\"  if the char exists
    pos = InStr(start, path, Chr$(92))
        
    If (pos <> 0) Then
        ' not at the last directory in the path string...
        directory = directory + Mid$(path, start, pos - start) + Chr$(92)
        If InStr(1, Mid$(path, start, pos - start), Chr$(58)) = 0 And Dir(directory, vbDirectory) = "" Then
           MkDir Mid$(directory, 1, Len(directory) - 1)
        End If
        ' call itself
        result = CreateDir(path)
    ElseIf (pos = 0) Then
        ' the last directory or the only in the path string
        directory = directory + Mid$(path, start, Len(path) - start + 1)
        MkDir Mid$(directory, 1, Len(directory))
        directory = ""
    End If
        
    ' success return true
    CreateDir = result
    
Exit Function

' if it gets here, an exception was thrown
' propogate the error to the calling function
errCreation:
    Err.Clear
    result = False
    CreateDir = result
        
End Function


