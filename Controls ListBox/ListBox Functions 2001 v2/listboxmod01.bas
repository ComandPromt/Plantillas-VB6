Attribute VB_Name = "listboxmod01"
Option Explicit
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Const LB_ADDSTRING& = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRINGEXACT& = &H1A2
Public Const LB_GETCOUNT& = &H18B
Public Const LB_GETCURSEL& = &H188
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN& = &H18A
Public Const LB_INSERTSTRING = &H181
Public Const LB_RESETCONTENT& = &H184
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_SETSEL = &H185
Public Sub xListKillDupes(listbox As listbox)
'Kills dublicite items in a listbox
        Dim Search1 As Long
        Dim Search2 As Long
        Dim KillDupe As Long
KillDupe = 0
For Search1& = 0 To listbox.ListCount - 1
For Search2& = Search1& + 1 To listbox.ListCount - 1
KillDupe = KillDupe + 1
If listbox.List(Search1&) = listbox.List(Search2&) Then
listbox.RemoveItem Search2&
Search2& = Search2& - 1
End If
Next Search2&
Next Search1&
End Sub

Public Function xListToTextString(listbox As listbox, InsertSeparator As String) As String
'Makes list a txt string

        Dim CurrentCount As Long, PrepString As String
For CurrentCount& = 0 To listbox.ListCount - 1
PrepString$ = PrepString$ & listbox.List(CurrentCount&) & InsertSeparator$
Next CurrentCount&
xListToTextString$ = Left(PrepString$, Len(PrepString$) - 2)
End Function
Public Sub xListCopy(SourceList As Long, DestinationList As Long)
'Copys a list to another
'Call ListCopy ("list1", "List2")
        Dim SourceCount As Long, OfCountForIndex As Long, FixedString As String
SourceCount& = SendMessageLong(SourceList&, LB_GETCOUNT, 0&, 0&)
Call SendMessageLong(DestinationList&, LB_RESETCONTENT, 0&, 0&)
If SourceCount& = 0& Then Exit Sub
For OfCountForIndex& = 0 To SourceCount& - 1
FixedString$ = String(250, 0)
Call SendMessageByString(SourceList&, LB_GETTEXT, OfCountForIndex&, FixedString$)
Call SendMessageByString(DestinationList&, LB_ADDSTRING, 0&, FixedString$)
Next OfCountForIndex&
End Sub

Public Function xListGetText(listbox As Long, index As Long) As String
        Dim ListText As String * 256
Call SendMessageByString(listbox&, LB_GETTEXT, index&, ListText$)
xListGetText$ = ListText$
End Function

Public Sub xListRemoveSelected(listbox As listbox)
        Dim ListCount As Long
ListCount& = listbox.ListCount
Do While ListCount& > 0&
ListCount& = ListCount& - 1
If listbox.Selected(ListCount&) = True Then
listbox.RemoveItem (ListCount&)
End If
Loop
End Sub
Public Sub xLoad2listboxes(Path As String, list1 As listbox, List2 As listbox)
'Loads Two list boxes
        Dim MyString As String, String1 As String, String2 As String
On Error Resume Next
Open Path$ For Input As #1
While Not EOF(1)
Input #1, MyString$
String1$ = Left(MyString$, InStr(MyString$, "*") - 1)
String2$ = Right(MyString$, Len(MyString$) - InStr(MyString$, "*"))
DoEvents
list1.AddItem String1$
List2.AddItem String2$
Wend
Close #1
End Sub
Public Function xListClickEvent()
'Have you ever wanted, on a listbox, that when a certain item is click, something
'happens, well, this is the coding for it
'Do not use this as in a module, but in the form, im just showing how its done.

'Private Sub List1_Click()
'If List1.List(List1.ListIndex) = "Source" Then
'MsgBox "You Click Source in List1"
'End If
'End Sub

End Function

Public Sub xSaveList(FileName As String, List As listbox)
    'self explanatory
    On Error Resume Next
    Dim lngSave As Long
    
    If FileName$ = "" Then Exit Sub
    
    Open FileName$ For Output As #1
        For lngSave& = 0 To List.ListCount - 1
            Print #1, List.List(lngSave&)
        Next lngSave&
    Close #1
End Sub
Public Sub xLoadList(FileName As String, List As listbox)
    'self explanatory
    Dim lstInput As String
    On Error Resume Next
    Open FileName$ For Input As #1
    While Not EOF(1)
        Input #1, lstInput$
        'DoEvents
        List.AddItem ReplaceText(lstInput$, "@aol.com", "")
    Wend
    Close #1
End Sub
Public Function ReplaceText(tMain As String, tFind As String, tReplace As String) As String
    'replaces a string within a larger string
    Dim iFind As Long, lString As String, rString As String, rText As String, tMain2 As String
    
    iFind& = InStr(1, LCase(tMain$), LCase(tFind$))
    If iFind& = 0& Then ReplaceText = tMain$: Exit Function
    
    Do
        DoEvents
        
        lString$ = Left(tMain$, iFind& - 1)
        rString$ = Mid(tMain$, iFind& + Len(tFind$), Len(tMain$) - (Len(lString$) + Len(tFind$)))
        tMain$ = lString$ + "" + tReplace$ + "" + rString$
        
        iFind& = InStr(iFind& + Len(tReplace$), LCase(tMain$), LCase(tFind$))
        If iFind& = 0& Then Exit Do
    Loop
    
    ReplaceText = tMain$
End Function
Public Sub SaveListBox(Directory As String, TheList As listbox)
    
    Dim savelist As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For savelist& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(savelist&)
    Next savelist&
    Close #1
End Sub
Public Sub Loadlistbox(Directory As String, TheList As listbox)
   
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        TheList.AddItem MyString$
    Wend
    Close #1
End Sub
Public Sub Load2Lists(ListSN As listbox, ListPW As listbox, Target As String)
    'self explanatory
    On Error Resume Next
    
    Dim lstInput As String, strSN As String, strPW As String
    
    If FileExists(Target$) = True Then
        Open Target$ For Input As #1
            While Not EOF(1) = True
                'DoEvents
                Input #1, lstInput$
                If InStr(1, lstInput$, "]-[") <> 0& And InStr(1, lstInput$, "=") <> 0& Then
                    lstInput$ = Mid(lstInput$, InStr(1, lstInput$, "]-[") + 3, Len(lstInput$) - 6)
                    strSN$ = Left(lstInput$, InStr(1, lstInput$, "=") - 1)
                    strPW$ = Right(lstInput$, Len(lstInput$) - InStr(1, lstInput$, "="))
                    If Trim(strSN$) <> "" And Trim(strPW$) <> "" Then
                        ListSN.AddItem Trim(strSN$)
                        ListPW.AddItem Trim(strPW$)
                    End If
                ElseIf InStr(1, lstInput$, ":") <> 0& Then
                    strSN$ = Left(lstInput$, InStr(1, lstInput$, ":") - 1)
                    strPW$ = Right(lstInput$, Len(lstInput$) - InStr(1, lstInput$, ":"))
                    If Trim(strSN$) <> "" And Trim(strPW$) <> "" Then
                        ListSN.AddItem Trim(strSN$)
                        ListPW.AddItem Trim(strPW$)
                    End If
                ElseIf InStr(1, lstInput$, "-") Then
                    strSN$ = Left(lstInput$, InStr(1, lstInput$, "-") - 1)
                    strPW$ = Right(lstInput$, Len(lstInput$) - InStr(1, lstInput$, "-"))
                    If Trim(strSN$) <> "" And Trim(strPW$) <> "" Then
                        ListSN.AddItem Trim(strSN$)
                        ListPW.AddItem Trim(strPW$)
                    End If
                ElseIf InStr(1, lstInput$, "=") Then
                    strSN$ = Left(lstInput$, InStr(1, lstInput$, "=") - 1)
                    strPW$ = Right(lstInput$, Len(lstInput$) - InStr(1, lstInput$, "="))
                    If Trim(strSN$) <> "" And Trim(strPW$) <> "" Then
                        ListSN.AddItem Trim(strSN$)
                        ListPW.AddItem Trim(strPW$)
                    End If
                ElseIf InStr(1, lstInput$, "·") Then
                    strSN$ = Left(lstInput$, InStr(1, lstInput$, "·") - 1)
                    strPW$ = Right(lstInput$, Len(lstInput$) - InStr(1, lstInput$, "·"))
                    If Trim(strSN$) <> "" And Trim(strPW$) <> "" Then
                        ListSN.AddItem Trim(strSN$)
                        ListPW.AddItem Trim(strPW$)
                    End If
                End If
            Wend
        Close #1
    End If
End Sub
Public Sub Save2Lists(ListSN As listbox, ListPW As listbox, Target As String)
    'self explanatory
    Dim sLong As Long
    On Error Resume Next
    
    Open Target$ For Output As #1
        For sLong& = 0 To ListSN.ListCount - 1
            Print #1, "" + ListSN.List(sLong&) + ":" + ListPW.List(sLong&) + ""
        Next sLong&
    Close #1
End Sub
Public Function FileExists(TheFileName As String) As Boolean
'Sees if the string(file) you specified exists
If Len(TheFileName$) = 0 Then
FileExists = False
Exit Function
End If
If Len(Dir$(TheFileName$)) Then
FileExists = True
Else
FileExists = False
End If
End Function
Public Function ListBoxCheckDup(List As listbox, Query As String) As Boolean
'knot n chichris
    If Query = "" Then Exit Function
    If Not TypeOf List Is listbox Then Exit Function
    Dim X As Long
    
    X = SendMessageByString(List.hwnd, LB_FINDSTRINGEXACT, 0&, Query)
    ListBoxCheckDup = IIf(X > -1, True, False)
End Function
Public Sub ClearListBoxes(frmTarget As Form)
    Dim i, j, ctrltarget


    For i = 0 To (frmTarget.Controls.Count - 1)
        Set ctrltarget = frmTarget.Controls(i)


        If TypeOf ctrltarget Is listbox Then
            ctrltarget.Clear
        End If
    Next i
End Sub


Public Sub AddHScroll(List As listbox)
    Dim i As Integer, intGreatestLen As Integer, lngGreatestWidth As Long
    'Find Longest Text in Listbox


    For i = 0 To List.ListCount - 1


        If Len(List.List(i)) > Len(List.List(intGreatestLen)) Then
            intGreatestLen = i
        End If
    Next i
    'Get Twips
    lngGreatestWidth = List.Parent.TextWidth(List.List(intGreatestLen) + Space(1))
    'Space(1) is used to prevent the last Ch
    '     aracter from being cut off
    'Convert to Pixels
    lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
    'Use api to add scrollbar
    SendMessage List.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0
    
End Sub
Sub SaveFormState(ByVal SourceForm As Form)
 Dim a As Long ' general purpose
 Dim B As Long
 Dim C As Long
 Dim FileName As String ' where to save to
 Dim FHandle As Long ' FileHandle
 ' error handling code
 On Error GoTo fError
 ' we create a filename based on the formname
 FileName = App.Path + "\" + SourceForm.Name + ".set"
 ' Get a filehandle
 FHandle = FreeFile()
 ' open the file
 #If DebugMode = 1 Then
  Debug.Print "--------------------------------------------------------->"
  Debug.Print "Saving Form State:" + SourceForm.Name
  Debug.Print "FileName=" + FileName
 #End If
 Open FileName For Output As FHandle
 ' loop through all controls
 ' first we save the type then the name
 For a = 0 To SourceForm.Controls.Count - 1
  #If DebugMode = 1 Then
   Debug.Print "Saving control:" + SourceForm.Controls(a).Name
  #End If
  ' if its textbox we save the .text property
  If TypeOf SourceForm.Controls(a) Is TextBox Then
   Print #FHandle, "TextBox"
   Print #FHandle, SourceForm.Controls(a).Name
   Print #FHandle, "StartText"
   Print #FHandle, SourceForm.Controls(a).Text
   Print #FHandle, "EndText"
    ' print a separator
   Print #FHandle, "|<->|"
  End If
  ' if its a checkbox we save the .value property
  If TypeOf SourceForm.Controls(a) Is CheckBox Then
   Print #FHandle, "CheckBox"
   Print #FHandle, SourceForm.Controls(a).Name
   Print #FHandle, Str(SourceForm.Controls(a).Value)
    ' print a separator
   Print #FHandle, "|<->|"
  End If
  ' if its a option button we save its value
  If TypeOf SourceForm.Controls(a) Is OptionButton Then
   Print #FHandle, "OptionButton"
   Print #FHandle, SourceForm.Controls(a).Name
   Print #FHandle, Str(SourceForm.Controls(a).Value)
    ' print a separator
   Print #FHandle, "|<->|"
  End If
  ' if its a listbox we save the .text and list contents
  If TypeOf SourceForm.Controls(a) Is listbox Then
   Print #FHandle, "ListBox"
   Print #FHandle, SourceForm.Controls(a).Name
   Print #FHandle, SourceForm.Controls(a).Text
   Print #FHandle, "StartList"
   For B = 0 To SourceForm.Controls(a).ListCount - 1
    Print #FHandle, SourceForm.Controls(a).List(B)
   Next B
   Print #FHandle, "EndList"
   ' save listindex
   Print #FHandle, CStr(SourceForm.Controls(a).ListIndex)
    ' print a separator
   Print #FHandle, "|<->|"
  End If
  ' if its a combobox, save .text and list items
  If TypeOf SourceForm.Controls(a) Is ComboBox Then
   Print #FHandle, "ComboBox"
   Print #FHandle, SourceForm.Controls(a).Name
   Print #FHandle, SourceForm.Controls(a).Text
   Print #FHandle, "StartList"
   For B = 0 To SourceForm.Controls(a).ListCount - 1
    Print #FHandle, SourceForm.Controls(a).List(B)
   Next B
   Print #FHandle, "EndList"
    ' print a separator
   Print #FHandle, "|<->|"
  End If
 Next a
' close file
 #If DebugMode = 1 Then
  Debug.Print "Closing File."
  Debug.Print "<----------------------------------------------------------"
 #End If
 Close #FHandle
 ' stop error handler
 On Error GoTo 0
 Exit Sub
fError: ' Simple error handler
 C = MsgBox("Error in SaveFormState. " + Err.Description + ", Number=" + CStr(Err.Number), vbAbortRetryIgnore)
 If C = vbIgnore Then Resume Next
 If C = vbRetry Then Resume
 ' else abort
End Sub
'@===========================================================================
' LoadFormState:
'  Loads the state of controls from file
'
'  Currently Supports: TextBox, CheckBox, OptionButton, Listbox, ComboBox
'=============================================================================
Sub LoadFormState(ByVal SourceForm As Form)
 Dim a As Long ' general purpose
 Dim B As Long
 Dim C As Long
 
 Dim txt As String ' general purpose
 Dim fData As String ' used to hold File Data
' these are variables used for controls data
 Dim cType As String ' Type of control
 Dim Cname As String ' Name of control
 Dim cNum As Integer ' number of control
' vars for the file
 Dim FileName As String ' where to save to
 Dim FHandle As Long ' FileHandle
 ' error handling code
 'On Error GoTo fError
 ' we create a filename based on the formname
 FileName = App.Path + "\" + SourceForm.Name + ".set"
 ' abort if file does not exist
 If Dir(FileName) = "" Then
  #If DebugMode = 1 Then
   Debug.Print "File Not found:" + FileName
  #End If
  Exit Sub
 End If
 ' Get a filehandle
 FHandle = FreeFile()
 ' open the file
 #If DebugMode = 1 Then
  Debug.Print "------------------------------------------------------>"
  Debug.Print "Loading FormState:" + SourceForm.Name
  Debug.Print "FileName:" + FileName
 #End If
 Open FileName For Input As FHandle
' go through file
 While Not EOF(FHandle)
  Line Input #FHandle, cType
  Line Input #FHandle, Cname
  ' Get control number
  cNum = -1
  For a = 0 To SourceForm.Controls.Count - 1
   If SourceForm.Controls(a).Name = Cname Then cNum = a
  Next a
  ' add some debug info if in debugmode
  #If DebugMode = 1 Then
   Debug.Print "Control Type=" + cType
   Debug.Print "Control Name=" + Cname
   Debug.Print "Control Number=" + CStr(cNum)
  #End If
  ' if we find control
  If Not cNum = -1 Then
   ' Depending on type of control, what data we get
   Select Case cType
   Case "TextBox"
    Line Input #FHandle, fData
    fData = "": txt = ""
    While Not fData = "EndText"
     If Not txt = "" Then txt = txt + vbCrLf
     txt = txt + fData
     Line Input #FHandle, fData
    Wend
    ' update control
    SourceForm.Controls(cNum).Text = txt
   Case "CheckBox"
    ' we get the value
    Line Input #FHandle, fData
    ' update control
    SourceForm.Controls(cNum).Value = fData
   Case "OptionButton"
    ' we get the value
    Line Input #FHandle, fData
    ' update control
    SourceForm.Controls(cNum).Value = fData
   Case "ListBox"
    ' clear listbox
    SourceForm.Controls(cNum).Clear
    ' get .text property
    Line Input #FHandle, fData
    SourceForm.Controls(cNum).Text = fData
    ' read past /startlist
    Line Input #FHandle, fData
    fData = "": txt = ""
    ' Get List
    While Not fData = "EndList"
     If Not fData = "" Then SourceForm.Controls(cNum).AddItem fData
     Line Input #FHandle, fData
    Wend
    ' get listindex
     Line Input #FHandle, fData
     SourceForm.Controls(cNum).ListIndex = Val(fData)
   Case "ComboBox"
    ' Clear combobox
    SourceForm.Controls(cNum).Clear
    ' Get Text
    Line Input #FHandle, fData
    SourceForm.Controls(cNum).Text = fData
    ' readpast /startlist
    Line Input #FHandle, fData
    fData = "": txt = ""
    ' get list
    While Not fData = "EndList"
     If Not fData = "" Then SourceForm.Controls(cNum).AddItem fData
     Line Input #FHandle, fData
    Wend
   End Select ' what type of control
  End If ' if we found control
  ' read till seperator
  fData = ""
  While Not fData = "|<->|"
   Line Input #FHandle, fData
  Wend
 Wend ' not end of File (EOF)
' close file
 #If DebugMode = 1 Then
  Debug.Print "Closing file.."
  Debug.Print "<------------------------------------------------------"
 #End If
 Close #FHandle
 Exit Sub
fError: ' Simple error handler
 C = MsgBox("Error in LoadFormState. " + Err.Description + ", Number=" + CStr(Err.Number), vbAbortRetryIgnore)
 If C = vbIgnore Then Resume Next
 If C = vbRetry Then Resume
 ' else abort
End Sub
Sub add1(listbox As listbox)
listbox.AddItem "test"
listbox.AddItem "test2"
listbox.AddItem "test3"
listbox.AddItem "test4"
listbox.AddItem "test5"
listbox.AddItem "test6"
End Sub
Sub add2(listbox As listbox, txt As String)
listbox.AddItem txt$
listbox.AddItem txt$
listbox.AddItem txt$
End Sub



