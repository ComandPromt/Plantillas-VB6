VERSION 5.00
Begin VB.Form frmReport 
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9435
   ControlBox      =   0   'False
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   9435
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   3000
      TabIndex        =   7
      Top             =   3120
      Width           =   2775
      Begin VB.TextBox txtStartDate 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Text            =   "Text1"
         ToolTipText     =   "Enter a start date for reporting"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtEndDate 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text1"
         ToolTipText     =   "Enter an end date for reporting"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Invoicing"
      Height          =   975
      Left            =   3120
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
      Begin VB.CommandButton btnInvoiceMark 
         Caption         =   "Mark Times Invoiced"
         Height          =   615
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Click to to mark as invoiced the times for the dates selected"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reporting"
      Height          =   975
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
      Begin VB.CommandButton btnInvoicing 
         Caption         =   "Invoicing Report"
         Height          =   615
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Click to see the un-invoiced hours for the dates selected"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8400
      TabIndex        =   4
      ToolTipText     =   "Click to exit"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblProgress 
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   4440
      Width           =   4095
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExit_Click()
  Unload Me
End Sub


Private Sub btnInvoiceMark_Click()
  If vbYes = MsgBox("Warning: If you mark a range of times as invoiced they will no longer appear " & Chr(10) & _
  "in Invoicing reports and users will no longer be able to edit them." & Chr(10) & Chr(10) & _
  "Are you sure you wish to mark this date range as Invoiced?", vbYesNo) Then
    If IsDate(Me.txtStartDate) = True Then
      If IsDate(Me.txtEndDate) = True Then
        reportingCode.markRangeInvoiced CDate(Me.txtStartDate), CDate(Me.txtEndDate)
      Else
        reportingCode.markRangeInvoiced CDate(Me.txtStartDate), CDate(Me.txtStartDate)
      End If
    Else
      MsgBox "Error: Start Date must be a valid date"
    End If
  End If
End Sub

Private Sub btnInvoicing_Click()
  If IsDate(Me.txtStartDate) = True Then
    If IsDate(Me.txtEndDate) = True Then
      reportingCode.getTimeReport CDate(Me.txtStartDate), CDate(Me.txtEndDate)
    Else
      reportingCode.getTimeReport CDate(Me.txtStartDate), CDate(Me.txtStartDate)
    End If
  End If
End Sub

Private Sub Form_Load()
  Me.txtEndDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
  Me.txtStartDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
End Sub
