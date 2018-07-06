VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#1.5#0"; "AgentCtl.dll"
Begin VB.UserControl AgentPro 
   BackColor       =   &H00000000&
   ClientHeight    =   264
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2064
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   264
   ScaleWidth      =   2064
   Begin AgentObjectsCtl.Agent ctlAgent 
      Left            =   240
      Top             =   480
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2052
   End
End
Attribute VB_Name = "AgentPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCharacters As clsCharacters

'Event Declarations:
Event ActivateInput(CharacterID As String) 'MappingInfo=ctlAgent,ctlAgent,-1,ActivateInput
Event BalloonHide(CharacterID As String) 'MappingInfo=ctlAgent,ctlAgent,-1,BalloonHide
Event BalloonShow(CharacterID As String) 'MappingInfo=ctlAgent,ctlAgent,-1,BalloonShow
Event Bookmark(BookmarkID As Long) 'MappingInfo=ctlAgent,ctlAgent,-1,Bookmark
Event Click(CharacterID As String, Button As Integer, Shift As Integer, x As Integer, y As Integer) 'MappingInfo=ctlAgent,ctlAgent,-1,Click
Event Command(UserInput As Object) 'MappingInfo=ctlAgent,ctlAgent,-1,Command
Event DblClick(CharacterID As String, Button As Integer, Shift As Integer, x As Integer, y As Integer) 'MappingInfo=ctlAgent,ctlAgent,-1,DblClick
Event DeactivateInput(CharacterID As String) 'MappingInfo=ctlAgent,ctlAgent,-1,DeactivateInput
Event DragComplete(CharacterID As String, Button As Integer, Shift As Integer, x As Integer, y As Integer) 'MappingInfo=ctlAgent,ctlAgent,-1,DragComplete
Event DragStart(CharacterID As String, Button As Integer, Shift As Integer, x As Integer, y As Integer) 'MappingInfo=ctlAgent,ctlAgent,-1,DragStart
Event Hide(CharacterID As String, Cause As Integer) 'MappingInfo=ctlAgent,ctlAgent,-1,Hide
Event IdleComplete(CharacterID As String) 'MappingInfo=ctlAgent,ctlAgent,-1,IdleComplete
Event IdleStart(CharacterID As String) 'MappingInfo=ctlAgent,ctlAgent,-1,IdleStart
Event RequestComplete(Request As Object) 'MappingInfo=ctlAgent,ctlAgent,-1,RequestComplete
Event RequestStart(Request As Object) 'MappingInfo=ctlAgent,ctlAgent,-1,RequestStart
Event Restart() 'MappingInfo=ctlAgent,ctlAgent,-1,Restart
Event Show(CharacterID As String, Cause As Integer) 'MappingInfo=ctlAgent,ctlAgent,-1,Show
Event Shutdown() 'MappingInfo=ctlAgent,ctlAgent,-1,Shutdown

Private Sub UserControl_Initialize()

    Set MyControl = Me
    Set mCharacters = New clsCharacters

End Sub

Public Property Get Characters() As clsCharacters
    Set Characters = mCharacters
End Property

Private Sub UserControl_InitProperties()
    If Ambient.UserMode = False Then
        UserControl.Size lblName.Width, lblName.Height
        lblName.Caption = UserControl.Ambient.DisplayName
    End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Size lblName.Width, lblName.Height
End Sub

Private Sub ctlAgent_ActivateInput(ByVal CharacterID As String)
    RaiseEvent ActivateInput(CharacterID)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ctlAgent,ctlAgent,-1,AudioOutput
Public Property Get AudioOutput() As IAgentCtlAudioObject
    Set AudioOutput = ctlAgent.AudioOutput
End Property

Private Sub ctlAgent_BalloonHide(ByVal CharacterID As String)
    RaiseEvent BalloonHide(CharacterID)
End Sub

Private Sub ctlAgent_BalloonShow(ByVal CharacterID As String)
    RaiseEvent BalloonShow(CharacterID)
End Sub

Private Sub ctlAgent_Bookmark(ByVal BookmarkID As Long)
    RaiseEvent Bookmark(BookmarkID)
End Sub

Private Sub ctlAgent_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    RaiseEvent Click(CharacterID, Button, Shift, x, y)
End Sub

Private Sub ctlAgent_Command(ByVal UserInput As Object)
    RaiseEvent Command(UserInput)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ctlAgent,ctlAgent,-1,CommandsWindow
Public Property Get CommandsWindow() As IAgentCtlCommandsWindow
    Set CommandsWindow = ctlAgent.CommandsWindow
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ctlAgent,ctlAgent,-1,Connected
Public Property Get Connected() As Boolean
    Connected = ctlAgent.Connected
End Property

Public Property Let Connected(ByVal New_Connected As Boolean)
    ctlAgent.Connected() = New_Connected
    PropertyChanged "Connected"
End Property

Private Sub ctlAgent_DblClick(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    RaiseEvent DblClick(CharacterID, Button, Shift, x, y)
End Sub

Private Sub ctlAgent_DeactivateInput(ByVal CharacterID As String)
    RaiseEvent DeactivateInput(CharacterID)
End Sub

Private Sub ctlAgent_DragComplete(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    RaiseEvent DragComplete(CharacterID, Button, Shift, x, y)
End Sub

Private Sub ctlAgent_DragStart(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    RaiseEvent DragStart(CharacterID, Button, Shift, x, y)
End Sub

Private Sub ctlAgent_Hide(ByVal CharacterID As String, ByVal Cause As Integer)
    RaiseEvent Hide(CharacterID, Cause)
End Sub

Private Sub ctlAgent_IdleComplete(ByVal CharacterID As String)
    RaiseEvent IdleComplete(CharacterID)
End Sub

Private Sub ctlAgent_IdleStart(ByVal CharacterID As String)
    RaiseEvent IdleStart(CharacterID)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ctlAgent,ctlAgent,-1,PropertySheet
Public Property Get PropertySheet() As IAgentCtlPropertySheet
    Set PropertySheet = ctlAgent.PropertySheet
End Property

Private Sub ctlAgent_RequestComplete(ByVal Request As Object)
    RaiseEvent RequestComplete(Request)
End Sub

Private Sub ctlAgent_RequestStart(ByVal Request As Object)
    RaiseEvent RequestStart(Request)
End Sub

Private Sub ctlAgent_Restart()
    RaiseEvent Restart
End Sub

Private Sub ctlAgent_Show(ByVal CharacterID As String, ByVal Cause As Integer)
    RaiseEvent Show(CharacterID, Cause)
End Sub

Private Sub ctlAgent_Shutdown()
    RaiseEvent Shutdown
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ctlAgent,ctlAgent,-1,SpeechInput
Public Property Get SpeechInput() As IAgentCtlSpeechInput
    Set SpeechInput = ctlAgent.SpeechInput
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ctlAgent,ctlAgent,-1,Suspended
Public Property Get Suspended() As Boolean
    Suspended = ctlAgent.Suspended
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ctlAgent,ctlAgent,-1,Size
Public Function Size(CharacterID As String, Width As Integer, Height As Integer) As Boolean
    Size = ctlAgent.Size(CharacterID, Width, Height)
End Function

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ctlAgent.Connected = PropBag.ReadProperty("Connected", False)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Connected", ctlAgent.Connected, False)
End Sub

Friend Property Get AgentObject() As Agent
    Set AgentObject = ctlAgent
End Property
