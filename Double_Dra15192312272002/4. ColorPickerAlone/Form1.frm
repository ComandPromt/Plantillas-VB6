VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "*Real-time* Color Picker"
   ClientHeight    =   3372
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   3924
   LinkTopic       =   "Form1"
   ScaleHeight     =   3372
   ScaleWidth      =   3924
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtForeColor 
      Height          =   315
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2520
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   315
      Left            =   3540
      TabIndex        =   0
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":0000
      Height          =   1032
      Left            =   120
      TabIndex        =   5
      Top             =   1260
      Width           =   3732
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":00DD
      Height          =   972
      Left            =   120
      TabIndex        =   4
      Top             =   180
      Width           =   3732
   End
   Begin VB.Label labForeColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   1140
      TabIndex        =   3
      Top             =   2520
      Width           =   312
   End
   Begin VB.Label Label1 
      Caption         =   "Fore Color:"
      Height          =   312
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
'   You are free to use this source as long as this copyright message
'     appears on your program's "About" dialog:
'
'   Outlook Bar Project
'   Copyright (c) 2002 Vlad Vissoultchev (wqw@geocities.com)
'
'=========================================================================
Option Explicit

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Any) As Long

Private Type UcsRgbQuad
    R                   As Byte
    G                   As Byte
    b                   As Byte
    a                   As Byte
End Type

Private m_clrFore           As OLE_COLOR

Private Property Get ForeColor_() As OLE_COLOR
    ForeColor_ = m_clrFore
End Property

Private Property Let ForeColor_(ByVal clrValue As OLE_COLOR)
    Dim rgbColor            As UcsRgbQuad
    
    m_clrFore = clrValue
    labForeColor.BackColor = m_clrFore
    OleTranslateColor m_clrFore, 0, rgbColor
    txtForeColor = "#" & pvHex(rgbColor.R) & pvHex(rgbColor.G) & pvHex(rgbColor.b)
End Property

Private Sub Command1_Click()
    Dim clrNew              As OLE_COLOR
    
    If frmColorPicker.Init(ForeColor_, clrNew) Then
        ForeColor_ = clrNew
    End If
End Sub

Private Sub Form_Load()
    ForeColor_ = labForeColor.BackColor
End Sub

Private Function pvHex(ByVal lValue As Long, Optional lCount As Long = 2) As String
    pvHex = Right(String(lCount, "0") & Hex(lValue), lCount)
End Function

