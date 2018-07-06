VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCharacters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

Public Function Add(ByVal CharacterID As String, ByVal CharacterFile As String) As clsCharacter
    'create a new object
    Dim objNewMember As clsCharacter
    Set objNewMember = New clsCharacter
    'Load the character of the underlying control
    objNewMember.Load CharacterID, CharacterFile
    mCol.Add objNewMember, CharacterID
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(CharacterID As String) As clsCharacter
    Set Item = mCol(CharacterID)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(CharacterID As String)
    mCol.Remove CharacterID
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

