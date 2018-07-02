VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmView 
   Caption         =   "Data Viewer"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtConView 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8916
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmView.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' Copyright © 2001 Gregory Kirk. All rights reserved.
'
' You have a royalty-free right to use, modify, reproduce and distribute the
' Application Files (and/or any modified version) in any way you find useful,
' provided that you agree that Gregory Kirk has no warranty, obligations or
' liability for any Application Files.
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' This is just a generic data dump viewer.
'-------------------------------------------------------------------------------

Option Explicit
Private Sub Form_Unload(Cancel As Integer)
txtConView.Text = ""
End Sub


