VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCharacterID As String
Private mCharacterFile As String
Private mCharacterObject As IAgentCtlCharacter

Property Let ID(newID As String)
    'Must unload and reload under new ID if changed
    If mCharacterObject Is Nothing Then
        mCharacterID = newID
    Else
        Set mCharacterObject = Nothing
        MyControl.AgentObject.Characters.Unload mCharacterID
        MyControl.AgentObject.Characters.Load newID, mCharacterFile
        mCharacterID = newID
        Set mCharacterObject = MyControl.AgentObject.Characters(mCharacterID)
    End If
End Property

Property Get ID() As String
    ID = mCharacterID
End Property

Public Function Speak(Optional Text As String, Optional URL As String) As IAgentCtlRequest
    If Not (Len(Trim(Text)) = 0 And Len(Trim(URL)) = 0) Then
        Set Speak = mCharacterObject.Speak(Text, URL)
    End If

End Function

Public Function MoveTo(x As Integer, y As Integer, Optional Speed As Integer = -1) As IAgentCtlRequest

    If Speed = -1 Then
        Set MoveTo = mCharacterObject.MoveTo(x, y)
    Else
        Set MoveTo = mCharacterObject.MoveTo(x, y, Speed)
    End If
End Function

Public Function GestureAt(x As Integer, y As Integer) As IAgentCtlRequest
    Set GestureAt = mCharacterObject.GestureAt(x, y)
End Function

Public Function Play(Animation As String) As IAgentCtlRequest
    Set Play = mCharacterObject.Play(Animation)
End Function

Public Function Load(CharacterID As String, LoadKey As String)
    MyControl.AgentObject.Characters.Load CharacterID, LoadKey
    mCharacterID = CharacterID
    mCharacterFile = LoadKey
    Set mCharacterObject = MyControl.AgentObject.Characters(CharacterID)
End Function

Private Sub Class_Terminate()
    MyControl.AgentObject.Characters.Unload mCharacterID
    Set mCharacterObject = Nothing
End Sub

Public Sub Greet()
    With mCharacterObject
        If .Visible = False Then .Show
        .Play "Wave"
        .Speak "Hello! | Hi there!"
        .Play "WaveReturn"
    End With
End Sub

Public Sub Dismiss()
    With mCharacterObject
        .Speak "Adios!"
        .Hide
    End With
End Sub

Public Function Show(Optional Fast As Boolean = False) As IAgentCtlRequest
    Set Show = mCharacterObject.Show(Fast)
End Function

Public Function Hide(Optional Fast As Boolean = False) As IAgentCtlRequest
    Set Hide = mCharacterObject.Hide(Fast)
End Function

Public Sub MoveRelative(hWnd As Long, xVariance As Long, yVariance As Long)
    Dim WindowRect As RECT, x As Long, y As Long
    GetWindowRect hWnd, WindowRect
    x = WindowRect.Left + xVariance
    y = WindowRect.Top + yVariance
    mCharacterObject.MoveTo x, y
End Sub

Public Sub GestureAtWindow(hWnd As Long)
    Dim WindowRect As RECT, x As Long, y As Long
    GetWindowRect hWnd, WindowRect
    x = Int((WindowRect.Right - WindowRect.Left) / 2)
    y = Int((WindowRect.Bottom - WindowRect.Top) / 2)
    mCharacterObject.GestureAt x, y
End Sub
