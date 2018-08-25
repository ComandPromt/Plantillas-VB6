Attribute VB_Name = "Module1"
'--------------------------------------------------------------------------------

'This code comes from part of my MAPIFunc module and mainly deals with sending rather than receiving email messages. The module was tested against Exchange on NT4 but should work on Win 95 - Outlook setups.

''
'' Created by E.Spencer (elliot@spnc.demon.co.uk) - This code is public domain.
''
Private Declare Function RegOpenKey Lib "AdvAPI32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, _
ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "AdvAPI32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, _
lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "AdvAPI32.dll" (ByVal hKey As Long) As Long
Private Const HKEY_CURRENT_USER = &H80000001
Public Enum SessMode
   StartSession
   StopSession
End Enum
Public Enum AddrType
   Primary
   CC
   BlindCC
End Enum
Public Enum AttachType
   DataFile
   OLEEmbedded
   OLEStatic
End Enum
Public SStatus, MStatus As String

' Call this function to start and stop MAPI sessions
' Example :- MyBool = AlterMailSession(Me, StartSession)
' MyBool = AlterMailSession(Me, StopSession)
' MyBool will be true if operation succeeded
' First parameter is reference to form that contains MAPI message / session controls
' Second parameter is the required session mode - stop or start.
Public Function AlterMailSession(ByRef FName As Form, Mode As SessMode) As Boolean
AlterMailSession = True
On Error GoTo SessError
If Mode = StartSession Then
   ' Get the default exchange profile name
   FName.MAPISession1.UserName = ReadRegistry(HKEY_CURRENT_USER, _
      "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\", "DefaultProfile")
   ' If session is already open return immediately
   If SStatus = "Open" Then Exit Function
   ' Set up profile - Default for exchange
   FName.MAPISession1.UserName = "MS Exchange Settings"
   FName.MAPISession1.SignOn ' Start mail session
   FName.MAPIMessages1.SessionID = FName.MAPISession1.SessionID ' Allocate session ID to Mail holder
   SStatus = "Open"
   MStatus = "Ready"
ElseIf Mode = StopSession Then
   ' If session is already closed return immediately
   If SStatus = "Closed" Then Exit Function
   FName.MAPISession1.SignOff ' End mail session
   FName.MAPIMessages1.SessionID = 0
   SStatus = "Closed"
   MStatus = "NotReady"
End If
Exit Function
SessError:
SStatus = "Closed"
MStatus = "NotReady"
AlterMailSession = False
End Function

' Call this function to start a new mail message
' Example :- MyBool = CreateMailMessage(Me, "Test Message", "Test Contents")
' MyBool will be true if operation succeeded
' First parameter is reference to form that contains MAPI message / session controls
' The second parameter is the message subject line.
' The third parameter is the message content text (embed Chr(13) for newlines)
Public Function
CreateMailMessage(ByRef FName As Form, Subject As String, Contents As String) As Boolean
CreateMailMessage = False
' If session is not open return immediately
If SStatus <> "Open" Then Exit Function
CreateMailMessage = True
On Error GoTo MessError
FName.MAPIMessages1.Compose ' Start new message composition
FName.MAPIMessages1.MsgSubject = Subject ' Insert message subject line
FName.MAPIMessages1.MsgNoteText = Contents & Chr(13) & " " ' Insert message text
MStatus = "Open"
Exit Function
MessError:
MStatus = "NotReady"
CreateMailMessage = False
End Function

' Call this function to abort a mail message
' Example :- MyBool = AbortMailMessage(Me)
' MyBool will be true if operation succeeded
' First parameter is reference to form that contains MAPI message / session controls
Public Function AbortMailMessage(ByRef FName As Form) As Boolean
AbortMailMessage = False
' If session is not open return immediately
If SStatus <> "Open" Then Exit Function
' If no current mail message then return immediately
If MStatus <> "Open" Then Exit Function
AbortMailMessage = True
On Error GoTo MessError
FName.MAPIMessages1.Delete (mapMessageDelete)
MStatus = "Ready"
Exit Function
MessError:
AbortMailMessage = False
End Function

' Call this function to send a complete mail message
' Example :- MyBool = SendMailMessage(Me)
' MyBool will be true if operation succeeded
' First parameter is reference to form that contains MAPI message / session controls
Public Function SendMailMessage(ByRef FName As Form) As Boolean
Dim Tries As Integer
SendMailMessage = False
' If session is not open return immediately
If SStatus <> "Open" Then Exit Function
' If no current mail message then return immediately
If MStatus <> "Open" Then Exit Function
SendMailMessage = True
On Error GoTo MessError
Retry:
FName.MAPIMessages1.Send
MStatus = "Ready"
Exit Function
MessError:
Tries = Tries + 1
If Tries < 10 Then GoTo Retry
SendMailMessage = False
End Function

' Call this function to save a complete mail message without sending it
' Example :- MyBool = SaveMailMessage(Me)
' MyBool will be true if operation succeeded
' First parameter is reference to form that contains MAPI message / session controls
Public Function SaveMailMessage(ByRef FName As Form) As Boolean
SaveMailMessage = False
' If session is not open return immediately
If SStatus <> "Open" Then Exit Function
' If no current mail message then return immediately
If MStatus <> "Open" Then Exit Function
SaveMailMessage = True
On Error GoTo MessError
FName.MAPIMessages1.Save
MStatus = "Ready"
Exit Function
MessError:
SaveMailMessage = False
End Function

' Call this function to address a mail message to a recipient
' Example :- MyBool = MailMessageTo(Me, "elliot spencer", Primary)
' MyBool will be true if operation succeeded. Supply display names from address book
' list - names will be resolved to addresses in the address book before being added to
' recipient list.
' First parameter is reference to form that contains MAPI message / session controls
' Second parameter is name of recipient (as displayed in address list)
' Third parameter is type of recipient
Public Function MailMessageTo(ByRef FName As Form, ToName As String, AddrMode As AddrType) As Boolean
MailMessageTo = False
' If session is not open return immediately
If SStatus <> "Open" Then Exit Function
' If no current mail message then return immediately
If MStatus <> "Open" Then Exit Function
MailMessageTo = True
On Error GoTo MessError
FName.MAPIMessages1.RecipIndex = FName.MAPIMessages1.RecipCount ' Update count of recipients
If AddrMode = Primary Then FName.MAPIMessages1.RecipType = 1 ' Set to primary recipient type
If AddrMode = CC Then FName.MAPIMessages1.RecipType = 2 ' Set to carbon copy type
If AddrMode = BlindCC Then FName.MAPIMessages1.RecipType = 3 ' Set to blind carbon copy type
FName.MAPIMessages1.RecipDisplayName = ToName ' Display name as provided
FName.MAPIMessages1.ResolveName ' Resolve display name to real address via address book
Exit Function
MessError:
MailMessageTo = False
End Function

' Call this function to address a mail message to a recipient
' Example :- MyBool = AddAttachment(Me, "Test File", "c:\test.txt", DataFile)
' MyBool will be true if operation succeeded.
' First parameter is reference to form that contains MAPI message / session controls
' Second parameter is name of recipient (as displayed in address list)
' Third parameter is type of recipient
Public Function AddAttachment(ByRef FName As Form, AName As String, APath As String, AttMode As AttachType) As Boolean
AddAttachment = False
' If session is not open return immediately
If SStatus <> "Open" Then Exit Function
' If no current mail message then return immediately
If MStatus <> "Open" Then Exit Function
AddAttachment = True
On Error GoTo MessError
FName.MAPIMessages1.AttachmentIndex = FName.MAPIMessages1.AttachmentCount ' Update count of attachments
If AttMode = DataFile Then FName.MAPIMessages1.AttachmentType = 0
If AttMode = OLEEmbedded Then FName.MAPIMessages1.AttachmentType = 1
If AttMode = OLEStatic Then FName.MAPIMessages1.AttachmentType = 2
FName.MAPIMessages1.AttachmentPosition = FName.MAPIMessages1.AttachmentIndex
FName.MAPIMessages1.AttachmentPathName = APath ' File or object path as provided
FName.MAPIMessages1.AttachmentName = AName ' File or object name as provided
Exit Function
MessError:
AddAttachment = False
End Function

' From my registry read module - just to get the default
' exchange user name (profile name)
'
Public Function ReadRegistry(ByVal Group As Long, ByVal Section As String, ByVal Key As String) As String
Dim lResult As Long, lKeyValue As Long, lDataTypeValue As Long, lValueLength As Long, sValue As String
On Error Resume Next
lResult = RegOpenKey(Group, Section, lKeyValue)
sValue = Space$(2048)
lValueLength = Len(sValue)
lResult = RegQueryValueEx(lKeyValue, Key, 0&, lDataTypeValue, sValue, lValueLength)
If (lResult = 0) And (Err.Number = 0) Then
   sValue = Left$(sValue, lValueLength - 1)
Else
   sValue = "Not Found"
End If
lResult = RegCloseKey(lKeyValue)
ReadRegistry = sValue
End Function

End Function


