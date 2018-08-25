VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmGeneratedCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generated Code"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGeneratedCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtCode 
      Height          =   4455
      Left            =   60
      TabIndex        =   3
      Top             =   420
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7858
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmGeneratedCode.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdToClipboard 
      Caption         =   "To Clipboard"
      Height          =   315
      Left            =   5940
      TabIndex        =   2
      Top             =   4980
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   7200
      TabIndex        =   1
      Top             =   4980
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Code generated by wizard:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   2115
   End
End
Attribute VB_Name = "frmgeneratedCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmGeneratedCode
'    Project    : EliteSpy
'
'    Description: Form for displaying the generated code
'
'    Author     : Andrea Batina
'    Modified   : 31/10/2001
'--------------------------------------------------------------------------------
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdToClipboard_Click()
    ' Clear clipboard
    Clipboard.Clear
    ' Set text on clipboard
    Clipboard.SetText txtCode.Text
End Sub
