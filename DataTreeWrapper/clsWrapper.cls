VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum FillBaseType
    dtwConnection = 1
    dtwRecordset = 2
End Enum

Dim m_Tree As TreeView

Dim m_DataEnvironmentCommands As DERuntimeObjects.Commands
Dim c_Command As ADODB.Command, c_RecordSet As ADODB.Recordset

Dim m_RecordSet As ADODB.Recordset

Dim m_MaxRecord As Long

Public RecordSetImage, RecordImage, KeyFieldImage, FieldImage

Public KeyFieldName  As String

Dim m_CommandKeyField(1 To 1000) As String

Public Graphical As Boolean

Public Property Let CommandKeyField(CommandID As Integer, str As String)
    m_CommandKeyField(CommandID) = str
End Property

Public Property Get CommandKeyField(CommandID As Integer) As String
    CommandKeyField = m_CommandKeyField(CommandID)
End Property

Public Property Let MaxRecord(lng As Long)
    m_MaxRecord = lng
End Property

Public Property Get MaxRecord() As Long
    MaxRecord = m_MaxRecord
End Property

Public Property Set DataEnvironmentCommands(DE As DERuntimeObjects.Commands)
    Set m_DataEnvironmentCommands = DE
    
    'ReDim CommandKeyField(1 To DE.Count) As String
    
End Property

Public Property Get DataEnvironmentCommands() As DERuntimeObjects.Commands
    Set DataEnvironmentCommands = m_DataEnvironmentCommands
End Property

Public Property Set Recordset(RS As ADODB.Recordset)
    Set m_RecordSet = RS
End Property

Public Property Get Recordset() As ADODB.Recordset
    Set Recordset = m_RecordSet
End Property

Public Property Set TheTree(Tree As TreeView)
    Set m_Tree = Tree
End Property

Public Property Get TheTree() As TreeView
    Set TheTree = m_Tree
End Property

Public Sub FillList(FillBase As FillBaseType, Optional DisplayRecordSetInfo As Boolean = True)

Dim TempNodes As MSComctlLib.Nodes

Set TempNodes = m_Tree.Nodes 'To backup the current nodes info

m_Tree.Nodes.Clear

Dim iRecord As Long
Dim i As Long

Select Case FillBase

    Case 1
        
        Dim iCommand As Integer
        
        With m_DataEnvironmentCommands
        
            For iCommand = 1 To .Count
                Set c_Command = .Item(iCommand)
                Set c_RecordSet = c_Command.Execute
                
                If Graphical = False Then
                    m_Tree.Nodes.Add , 4, "command" & iCommand, c_Command.CommandText
                Else
                    m_Tree.Nodes.Add , 4, "command" & iCommand, c_Command.CommandText, RecordSetImage
                End If
                
                With c_RecordSet
                    
                    '**********[ Start Filling RecordSet Info]*************************
                    If DisplayRecordSetInfo = True Then
                        
                        If Graphical = False Then
                            m_Tree.Nodes.Add "command" & iCommand, 4, "rsinfo" & iCommand, "Record Set Information"
                        Else
                            m_Tree.Nodes.Add "command" & iCommand, 4, "rsinfo" & iCommand, "Record Set Information", RecordSetImage
                        End If
                        
                        For i = 0 To c_RecordSet.Properties.Count - 1
                                
                            If Graphical = False Then
                                m_Tree.Nodes.Add "rsinfo" & iCommand, 4, "property" & iCommand & "|" & i, .Properties(i).Name & "=" & .Properties(i).Value
                            Else
                                m_Tree.Nodes.Add "rsinfo" & iCommand, 4, "property" & iCommand & "|" & i, .Properties(i).Name & "=" & .Properties(i).Value, RecordImage
                            End If
                                
                        Next i
                        
                    End If
                    '*******************************************************************
                    
                    iRecord = -1
                    
                    .MoveFirst
                    
                    Do Until (.EOF = True) Or (iRecord > MaxRecord)
                        DoEvents
                        
                        iRecord = iRecord + 1
                        
                        If Graphical = False Then
                            m_Tree.Nodes.Add "command" & iCommand, 4, "key" & iRecord & "|" & iCommand, .Fields(CommandKeyField(iCommand)).Name & ": " & .Fields(CommandKeyField(iCommand)).Value
                            
                            For i = 0 To .Fields.Count - 1
                                m_Tree.Nodes.Add "key" & iRecord & "|" & iCommand, 4, .Fields(i).Name & iRecord, .Fields(i).Name & ": " & .Fields(i).Value
                            Next i
                            
                        Else
                            m_Tree.Nodes.Add "command" & iCommand, 4, "key" & iRecord & "|" & iCommand, .Fields(CommandKeyField(iCommand)).Name & ": " & .Fields(CommandKeyField(iCommand)).Value, RecordImage
                            
                            For i = 0 To .Fields.Count - 1
                                If .Fields(i).Name <> CommandKeyField(iCommand) Then
                                    m_Tree.Nodes.Add "key" & iRecord & "|" & iCommand, 4, .Fields(i).Name & iRecord & "|" & iCommand, .Fields(i).Name & ": " & .Fields(i).Value, FieldImage
                                Else
                                    m_Tree.Nodes.Add "key" & iRecord & "|" & iCommand, 4, .Fields(i).Name & iRecord & "|" & iCommand, .Fields(i).Name & ": " & .Fields(i).Value, KeyFieldImage
                                End If
                            Next i
                            
                        End If
                        
                        .MoveNext
                        
                    Loop
                                    
                End With
                        
                        
            Next iCommand
        
        End With
        
    Case 2
    
        With m_RecordSet
            
            '**********[ Start Filling RecordSet Info]*************************
            If DisplayRecordSetInfo = True Then
                
                If Graphical = False Then
                    m_Tree.Nodes.Add "command" & iCommand, 4, "rsinfo", "Record Set Information"
                Else
                    m_Tree.Nodes.Add "command" & iCommand, 4, "rsinfo", "Record Set Information", RecordSetImage
                End If
                
                For i = 0 To m_RecordSet.Properties.Count - 1
                        
                    If Graphical = False Then
                        m_Tree.Nodes.Add "rsinfo", 4, "property" & i, .Properties(i).Name & "=" & .Properties(i).Value
                    Else
                        m_Tree.Nodes.Add "rsinfo", 4, "property" & i, .Properties(i).Name & "=" & .Properties(i).Value, RecordImage
                    End If
                        
                Next i
                
            End If
            '*******************************************************************
            
            iRecord = -1
            
            .MoveFirst
            
            Do Until (.EOF = True) Or (iRecord > MaxRecord)
                DoEvents
                
                iRecord = iRecord + 1
                
                If Graphical = False Then
                    m_Tree.Nodes.Add , 0, "key" & iRecord, .Fields(KeyFieldName).Name & ": " & .Fields(KeyFieldName).Value
                    
                    For i = 0 To .Fields.Count - 1
                        m_Tree.Nodes.Add "key" & iRecord, 4, .Fields(i).Name & iRecord, .Fields(i).Name & ": " & .Fields(i).Value
                    Next i
                    
                Else
                    m_Tree.Nodes.Add , 0, "key" & iRecord, .Fields(KeyFieldName).Name & ": " & .Fields(KeyFieldName).Value, RecordImage
                    
                    For i = 0 To .Fields.Count - 1
                        If .Fields(i).Name <> KeyFieldName Then
                            m_Tree.Nodes.Add "key" & iRecord, 4, .Fields(i).Name & iRecord, .Fields(i).Name & ": " & .Fields(i).Value, FieldImage
                        Else
                            m_Tree.Nodes.Add "key" & iRecord, 4, .Fields(i).Name & iRecord, .Fields(i).Name & ": " & .Fields(i).Value, KeyFieldImage
                        End If
                    Next i
                    
                End If
                
                .MoveNext
                
            Loop
                            
        End With
    
End Select

End Sub

Private Sub Class_Initialize()
MaxRecord = 500
End Sub
