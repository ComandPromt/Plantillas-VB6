VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUndoElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive
'*********************************************************************
' This is a heavily modified version of
'
' Multi Undo Class for RichTextBox v1.02
' Copyright ©2002 by Sebastian Thomschke, All Rights Reserved.
' http://www.sebthom.de
' You can download Sebastians classes at:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=34094
'
'
' Reasons for modifying:
' Sebastian developed these classes for a PHP editor and
' as such had no need to support large file sizes.
' As this is a text editor, it did not need the full undo
' functionality provided by Sebastians classes - just plain
' text here. With that in mind I have rewritten the majority
' of the classes to better suit these needs. The system is
' now quite different and much faster. Optimised the Undo/Redo stacks
' to only hold the actual undo or redo, thus avoiding parsing the
' entire file, which can cause an unacceptable lag when undoing/redoing
' on larger files. The structure is in place for full Rich Text editing
' support, though it is not implemented here.
'*********************************************************************

'Variables to hold Undo/Redo data
Public SelStart As Long
Public SelUndoLength As Long
Public SelRedoLength As Long
Public SelUndoText As String
Public SelRedoText As String
Public UndoTabPosition As Long
Public RedoTabPosition As Long
Public Action As Boolean

