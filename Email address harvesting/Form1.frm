VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Get_Email_Addresses 
   AutoRedraw      =   -1  'True
   Caption         =   "Get All Email Addresses"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5850
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Help"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Use Access DB-testing only"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.ProgressBar MessageProgress 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar FolderProgress 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process Sent Item Subfolders"
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Message Progress"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Folder Progress"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   3615
   End
End
Attribute VB_Name = "Get_Email_Addresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public MyOlApp As New Outlook.Application
 Public MyOlMessage As Outlook.MailItem
 Public MyOlSpace As Outlook.NameSpace
 Public MyFolder As Outlook.MAPIFolder
 Public MyOlMessageFolder As Outlook.MAPIFolder
 Public MoveToFolder As Outlook.MAPIFolder
 
 Public MsgCount, MyItem As Integer
 Public MyText As String
 Public currentmessage As Integer
 
Dim db As DAO.Database
Dim rsMsg As DAO.Recordset
Dim wrkODBC As DAO.Workspace

Public genie As IAgentCtlCharacterEx

Private Sub Command1_Click()
    Call ImportMessages("Sent Items", "Mailbox - Dynamic Tools Support Internet")
    Get_Email_Addresses.Hide
    Unload Get_Email_Addresses
End Sub


Private Sub Form_Load()
Get_Email_Addresses.Left = 3500
Get_Email_Addresses.Top = 2000
End Sub


'======================================================================
'FUNCTION: ParseRecipients
'
'Purpose: Check a MAPI message for a specific type of recipient and
'         return a semicolon delimited list of recipients. For
'         instance, if this function is called using the MapiTo
'         constant, this function will return a semicolon delimited
'         list of all recipients on the 'TO' line of the message.
'======================================================================
Function ParseRecipients(objMessage As Outlook.MailItem, ByRef displayname As String)
    Dim RecipientCount As Long
    Dim Recipient As Object
    Dim TheSender As Object
    Dim ReturnString As String
    Dim EmailName As String
    Dim result As String
    Dim messagecopy As Outlook.MailItem
    
    RecipientCount = objMessage.Recipients.Count
    If RecipientCount = 0 Then Exit Function
    Set Recipient = objMessage.Recipients(RecipientCount)
    If RecipientCount > 0 Then
        ReturnString = objMessage.Recipients(1).Address
    End If
    'We don't want to use any of the Tools emails as the recipient
    result = UCase((StripQuote(ReturnString)))
    
    If result = UCase("/O=GPS/OU=GPSDNS/CN=RESOURCES/CN=TEAMS/CN=DYNTOOLS") _
    Or result = UCase("dexsupport@gps.com") _
    Or result = UCase("dexsupport@GreatPlains.com") _
    Or result = UCase("dexsuprt@gps.com") _
    Or result = UCase("dexsuprt@GreatPlains.com") _
    Or result = UCase("dexterity_support@gps.com") _
    Or result = UCase("dexterity_support@GreatPlains.com") _
    Or result = UCase("dyntools@GPS.com") _
    Or result = UCase("dyntools@gpsdns.gps.com") _
    Or result = UCase("dyntools@GreatPlains.com") _
    Or result = UCase("tdexteri@gps.com") _
    Or result = UCase("tdexteri@GreatPlains.com") Then
        'if we see that it's to us, make a reply to we can get the senders
        'email address.  Have to do it this way because there is no way possible
        'to get the senders actual email address.
        'so make a copy by replying and then get the recipients (was the senders)
        'actual email address.  Delete it when finished.
        Set messagecopy = objMessage.Reply
        ReturnString = StripQuote(messagecopy.Recipients(1).AddressEntry.Address)
        displayname = messagecopy.To
        messagecopy.Delete
        GoTo mylabel
    End If

    For RecipientCount = 1 To objMessage.Recipients.Count
       
        Set Recipient = objMessage.Recipients(RecipientCount)
    
        EmailName = StripQuote(Recipient.Name)
        If displayname = "" Then
            displayname = EmailName
        End If
        If EmailName = ReturnString Then
            ReturnString = StripQuote(Recipient.Address)
        End If
    Next
mylabel:
    If Len(ReturnString) > 0 Then
        ReturnString = Left(Trim(ReturnString), Len(ReturnString))
        ParseRecipients = ReturnString
    Else
        ParseRecipients = Null
    End If
End Function


'======================================================================
'SUB: WriteMessage
'
'Purpose: Adds message information to fields in the table through the
'         the recordset opened in the ImportMessages Sub. This

'         procedure is called from the RetrieveMessage Sub when it is
'         time to write information to the table.
'======================================================================

Sub WriteMessage(objMessage As Object, FolderName As String, _
                 InfoStore As String)
    Dim RetVal
    Dim iString As String
    Dim displayname  As String
    
    On Error GoTo myerror
    txt = "select * from Contacts where Email = '" & ParseRecipients(objMessage, displayname) & "'"
    
    Set rsMsg = db.OpenRecordset(txt, dbOpenDynamic, dbExecDirect, dbOptimistic)
    
    On Error Resume Next
    rsMsg.MoveFirst
  
    If rsMsg.EOF Then
        
        With rsMsg
            .AddNew
            !CompanyName = FolderName
            !email = ParseRecipients(objMessage, displayname)
            If rsMsg!email = "" Then Exit Sub
            !displayname = displayname
            rsMsg.Update
        End With
    Else
        'never update the email address to new folder
        'With rsMsg
            '.Edit
            '!CompanyName = FolderName
            '!email = ParseRecipients(objMessage, displayname)
            '!displayname = displayname
            'rsMsg.Update
        'End With
    End If
myerror:
End Sub

'======================================================================
'SUB: RetrieveMessage
'
'Purpose: Loop through the Messages collection of each Folder of the

'         specified information store(s) and calls the WriteMessage Sub
'         to write individual messages to the table. This procedure is
'         called by the ImportMessages Sub.
'======================================================================

Sub RetrieveMessage(objInfoStore As Object, FolderName As Variant)
    Dim objFoldersColl As Object, objFolder As Object
    Dim objMessage As Object, objMessageColl As Object
    Dim olTarget As Outlook.MAPIFolder
    Dim MessageCount As Long
    Dim currentmessagecount As Long
    
    Dim x As Long, foldercount As Long

    'Set a Variable equal to the Folders Collection of the InfoStore's
    'Top Level Folder. (RootFolder)
    Set objFoldersColl = objInfoStore.RootFolder.Folders
    With objFoldersColl
        'Set objFolder = .GetFirst
        Set objFolder = objFoldersColl.Item("Sent Items")
        'Loop through each folder and determine if we're looking for a
        'specific folder from which we're importing messages, or all
        'folders.
        Do While Not objFolder Is Nothing

            If IsMissing(FolderName) Then
                Set objMessageColl = objFolder.Messages
                With objMessageColl
                    Set objMessage = .GetFirst
                    Do While Not objMessage Is Nothing
                        Call WriteMessage(objMessage, objFolder.Name, _
                                          objInfoStore.Name)
                        Set objMessage = .GetNext

                    Loop
                End With
                Set objFolder = .GetNext
            Else
                If objFolder.Name = FolderName Then
                    Set MyOlSpace = MyOlApp.GetNamespace("MAPI")
                    Get_Email_Addresses.Show
                    MsgBox "Select the folder to be processed.  Normally the Dyntools Sent Items"
                    Set MyFolder = MyOlSpace.PickFolder
                    Set olTarget = MyFolder
                    
                    foldercount = MyFolder.Folders.Count
                    FolderProgress.Min = 0
                    FolderProgress.Max = foldercount
                    For x = 1 To foldercount
                        'update the progress on folder
                        FolderProgress.Value = x
                        
                        Set olTarget = MyFolder.Folders.Item(x)
                        If olTarget.Name = "General" Then
                            'don't catalog anything in the General Box
                            GoTo continueHere
                            'x = x + 1
                            'olTarget = MyFolder.Folders.Item(x)
                            
                        End If
                        Set objMessageColl = olTarget.Items
                        MessageCount = olTarget.Items.Count
                        MessageProgress.Min = 0
                        If MessageCount = 0 Then
                            MessageCount = MessageCount + 1
                        End If
                        MessageProgress.Max = MessageCount
                        currentmessagecount = 0
                        With objMessageColl
                            Set objMessage = .GetFirst
                            Do While Not objMessage Is Nothing
                                currentmessagecount = currentmessagecount + 1
                                'update progress control on messages
                                MessageProgress.Value = currentmessagecount
                                Call WriteMessage(objMessage, _
                                    olTarget.Name, objInfoStore.Name)
                                Set objMessage = .GetNext
                            Loop
                        End With
continueHere:
                    
                    Next
                    Exit Do
                Else
                    Set objFolder = .GetNext
                End If
            End If
        Loop
    End With
End Sub


'======================================================================

'SUB: ImportMessage
'
'Purpose: Opens a MAPI session through OLE automation and opens a
'         recordset based on the Messages table. Then, this procedure
'         checks to see if it needs to import messages from top level
'         folders in ALL information stores, or just a specific
'         information store. Based upon this, the procedure will call
'         the RetrieveMessage sub for the specified information stores.
'======================================================================

Sub ImportMessages(Optional FolderName As Variant, _
                   Optional InfoStoreName As Variant)
    Dim objMapi As Object
    Dim objFoldersColl As Object
    Dim objInfoStore As Object
    Dim RetVal
    Dim foldercount As Integer
    
    'DoCmd.Hourglass True
    'On Error GoTo helpme
    strConnect = "ODBC;DSN=OutlookDB;UID=sa;PWD="

    If Check1 Then
        Set db = OpenDatabase("d:\outlook.mdb")
    Else
        If wrkODBC Is Nothing Then
            Set wrkODBC = DBEngine.CreateWorkspace("ODBC", "", "", dbUseODBC)
            Set db = wrkODBC.OpenDatabase("SQL Dynamics", dbDriverComplete, False, strConnect)
        End If
    End If
    wrkODBC.BeginTrans
    
    'Set rsMsg = db.OpenRecordset("Messages", dbOpenDynaset)
    

    'RetVal = SysCmd(acSysCmdSetStatus, "Establishing MAPI Session...")

    Set objMapi = CreateObject("Mapi.Session")
    'RetVal = SysCmd(acSysCmdSetStatus, "Logging on to MAPI Session...")

'In the following line, replace the ProfileName argument with a valid
'profile. If you omit the ProfileName argument, Microsoft Exchange will
'prompt you for your profile.

    objMapi.Logon ProfileName:=""

    'Loop through each InfoStore in the MAPI session and determine if
    'we should read in messages from ALL InfoStores or just a specified

    'InfoStore. InfoStores include a user's personal store files
    '(.PST Files), Network stores, and Public Folders.
       For Each objInfoStore In objMapi.InfoStores
            If Not IsMissing(InfoStoreName) Then
                If objInfoStore.Name = InfoStoreName Then
                    Call RetrieveMessage(objInfoStore, FolderName)
                    Exit For
                End If
              
            Else
                Call RetrieveMessage(objInfoStore, FolderName)
                Exit For
            End If
        Next

    objMapi.Logoff  ' Log out of the MAPI session.
    'Commit the Transactions
    wrkODBC.CommitTrans
    
    Set objMapi = Nothing
    db.Close  ' Close the Database.
    Set db = Nothing
    If wrkODBC Is Nothing Then
        Set wrkODBC = Nothing
    End If
    'DoCmd.Hourglass False
    'RetVal = SysCmd(acSysCmdClearStatus)
    Exit Sub
helpme:
    Call myerror

End Sub


Private Function StripQuote(instring As String) As String
    Dim x As Integer
    Dim i As Integer
    
    For x = 1 To Len(instring)
        If Mid(instring, x, 1) = ";" Or Mid(instring, x, 1) = "," Then
            Exit Function
        End If
        If Mid(instring, x, 1) <> "'" Then
            StripQuote = StripQuote & Mid(instring, x, 1)
        End If
    Next
End Function


Private Sub Command2_Click()
    If Command2.Caption = "Help" Then
        Command2.Caption = "Done"
             Set Agent1 = CreateObject("Agent.Control.1")
             Agent1.Connected = True
             Agent1.Characters.Load "Genie", "Genie.Acs"
             Set genie = Agent1.Characters("Genie")
             genie.Show
             genie.Top = 110
             genie.Left = 620


             genie.Speak ("Welcome to the Tools Outlook email harvesting program.")
             genie.Play ("Greet")
             genie.Play ("RestPose")
             genie.Play ("Explain")
             genie.Speak ("My mission today is to find all the email addresses of stored emails.")
             genie.Speak ("Emails to/from the tools team in each subfolder will be stored")
             genie.Speak ("on INTLDEV2 in a SQL Database.  Make sure that you have a SQL Driver DSN named OutlookDB.")
             genie.Speak ("It needs the Initial Database set to OutlookDB as well.")
             
             genie.Play ("RestPose")
    
             genie.Speak ("Press the 'Process Sent Items Subfolders' button to start.")
             genie.Speak ("First select your profile, the default profile will probably work fine.")
             genie.Speak ("Then choose the DynamicTools Sent Items folder.")
             genie.Speak ("And then Poof!")
             genie.Play ("DoMagic1")
             genie.Play ("DoMagic2")
             genie.Play ("RestPose")
             genie.Speak ("The new email addresses are now in the SQL database ready to be used the next time FileIt.exe is run.")
             genie.Play ("RestPose")
             genie.Play ("RestPose")
             genie.Play ("Idle3_1")
             genie.Play ("Idle3_2")
             While Command2.Caption = "Done"
                DoEvents
             Wend
        
    Else
        Command2.Caption = "Help"
        genie.Stop
        genie.Hide
    End If
    

End Sub

Private Sub myerror()
    If Command2.Caption = "Help" Then
        Command2.Caption = "Done"
             Set Agent1 = CreateObject("Agent.Control.1")
             Agent1.Connected = True
             Agent1.Characters.Load "Genie", "Genie.Acs"
             Set genie = Agent1.Characters("Genie")
             genie.Show
             genie.Top = 110
             genie.Left = 620
             genie.Speak ("You are here because the connection to INTLDEV2 could not be reached")
             genie.Speak ("The most likely cause is not having the DSN setup.  Make sure that you have a SQL Driver DSN named OutlookDB.")
             genie.Speak ("It needs the Initial Database set to OutlookDB as well.")
             
             While Command2.Caption = "Done"
                DoEvents
             Wend
        
    Else
        Command2.Caption = "Help"
        genie.Stop
        genie.Hide
    End If

End Sub
