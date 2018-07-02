Attribute VB_Name = "basMapping"
Option Explicit
'
'  Routines using the clsMapping class.
'
'  Saves/loades 1 or more mappings from/to
'  a file:

'  SaveMapping  - Saves a mapping to a file.
'  LoadMapping  - Loads a mapping from a file.
'
'
'  The following expect an open file handle:
'
'  ReadMapping  - Reads a mapping from a file.
'  WriteMapping - Writes a mapping to a file.
'
'
'  Saves or loads a mapping from the registry.
'
'  SaveSettingMapping
'  LoadSettingMapping
'
'  There are some limitations on the types of data
'  that can be saved and loaded.
'    - Variant types vbObject, vbDataObject, and vbError
'      can not be saved.
'

Private Const DEFAULT_FILE_KEY = 1347436877  'In asci this is "MAPP"
Private Const BASE_ERROR = 1000

Private vNull As Variant
Public Sub SaveSettingsMapping(m As clsMapping, appname As Variant, Optional Key As Variant, Optional clearsetting As Variant)
   Dim i As Integer
   Dim AppN As String
   Dim k As String
   
   If IsMissing(appname) Then
      AppN = App.ProductName
   Else
      AppN = CStr(appname)
   End If
   
   If IsMissing(Key) Then
      k = "Settings"
   Else
      k = CStr(Key)
   End If

   If Not IsMissing(clearsetting) Then
      If clearsetting Then
         On Error Resume Next
         DeleteSetting AppN, k
      End If
   End If
   
   For i = 1 To m.Count
      SaveSetting AppN, k, m.Key(i), m.Item(i)
   Next i
End Sub

Public Sub LoadSettingsMapping(m As clsMapping, appname As Variant, Optional Key As Variant)
   Dim AppN As String
   Dim k As String
   Dim v As Variant
   Dim i As Integer
   
   If IsMissing(appname) Then
      AppN = App.ProductName
   Else
      AppN = CStr(appname)
   End If
   
   If IsMissing(Key) Then
      k = "Settings"
   Else
      k = CStr(Key)
   End If
   
   v = GetAllSettings(AppN, k)
   If Not IsEmpty(v) Then
      For i = LBound(v, 1) To UBound(v, 1)
         m.Item(v(i, 0)) = v(i, 1)
      Next i
   End If
End Sub


'
'  If keynum = 0 then use the default keynum
'  If Keynum = -1 then ignore any keynum
'
Public Sub LoadMapping(filename As String, KeyNum As Long, ParamArray m() As Variant)
   Dim iErr As Integer
   Dim sErr As String
   Dim fh As Integer   ' File Handle
   Dim l As Long
   Dim i As Long
   
   On Error GoTo ErrorHandler
   
   fh = FreeFile(0)
   Open filename For Binary Access Read Lock Read Write As fh
   
   Get fh, , i
   
   If KeyNum = 0 Then
      l = CLng(DEFAULT_FILE_KEY)
      If l <> i Then
         Close fh
         
         On Error GoTo 0
         Err.Raise BASE_ERROR + 1, "LoadMapping", "File is corrupt or of an unknown format."
         
         Exit Sub
      End If
   ElseIf KeyNum = -1 Then
      ' do nothing
   Else
      l = CLng(KeyNum)
      If l <> 0 Then
         If l <> i Then
            Close fh
            
            On Error GoTo 0
            Err.Raise BASE_ERROR + 1, "LoadMapping", "File is corrupt or of an unknown format."
            
            Exit Sub
         End If
      End If
   End If
 
   For i = 0 To UBound(m)
      ReadMapping fh, m(i)
   Next i
   
   Close fh
   Exit Sub
   
   ' Do error handleing to make sure the file is
   ' closed, then pass the error to the main
   ' program
ErrorHandler:
   iErr = Err
   sErr = Err.Description
   
   On Error Resume Next
   Close fh
   
   On Error GoTo 0
   Err.Raise iErr, "SaveMapping", sErr
End Sub
Public Sub ReadMapping(FileHandle As Integer, m As Variant)
   Dim l As Long
   Dim k As Variant
   Dim v As Variant
   Dim i As Long
   
   Get FileHandle, , l
      
   For i = 1 To l
      Get FileHandle, , k
      Get FileHandle, , v
      
      If Not IsNull(k) Then
         m.Item(k) = v
      End If
   Next
End Sub


'
'  Uses the default KeyNum if KeyNum = 0
'
Public Sub SaveMapping(filename As String, ByVal KeyNum As Long, ParamArray m() As Variant)
   Dim iErr As Integer
   Dim sErr As String
   Dim fh As Integer   ' File Handle
   Dim l As Long
   Dim e As Boolean
   Dim e2 As Boolean
   Dim i As Long
   
   On Error GoTo ErrorHandler
   
   e2 = False
   fh = FreeFile(0)
   
   On Error Resume Next
   Kill filename
   
   On Error GoTo ErrorHandler
   Open filename For Binary Access Write Lock Read Write As fh
   
   If KeyNum = 0 Then
      l = CLng(DEFAULT_FILE_KEY)
      Put fh, , l
   Else
      l = CLng(KeyNum)
      Put fh, , l
   End If
   
   For i = 0 To UBound(m)
      e = WriteMapping(fh, m(i))
      If Not e Then
         e2 = True
      End If
   Next i
   
   Close fh
   
   On Error GoTo 0

   If e2 Then Err.Raise BASE_ERROR, "SaveMapping", _
      "All data was not of a valid type.  Some data may not have been saved."
   
   Exit Sub
   
   
   ' Do error handleing to make sure the file is
   ' closed, then pass the error to the main
   ' program
ErrorHandler:
   iErr = Err
   sErr = Err.Description
   
   On Error Resume Next
   Close fh
   
   On Error GoTo 0
   Err.Raise iErr, "SaveMapping", sErr
End Sub

'
'  Returns 1 if the item was not of a valid type,
'  and 0 if it was.
'
Private Function WriteItem(ByVal FileHandle As Integer, v As Variant) As Long
   Select Case VarType(v) And Not vbArray
      Case vbError, vbDataObject, vbObject:
         Put FileHandle, , vNull
         WriteItem = 1
         
      Case Else:
         Put FileHandle, , v
         WriteItem = 0
   End Select
End Function
'
'  Writes a mapping to the file associated with the
'  handle 'FileHandle'.
'
'  Returns False if some data was not written because
'  it was not of a type that could be saved.
'
Public Function WriteMapping(ByVal FileHandle As Integer, m As Variant) As Boolean
   Dim l As Long
   Dim i As Long
   Dim k As Variant
   Dim e As Long
   
   e = 0
   l = m.Count
   vNull = Null
   
   Put FileHandle, , l
   
   For i = 1 To l
      k = m.Key(i)
      
      e = e + WriteItem(FileHandle, k)
      e = e + WriteItem(FileHandle, m.Item(k))
   Next i
   
   WriteMapping = IIf(e > 0, False, True)
End Function
