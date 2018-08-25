VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form MsChartFrm 
   Caption         =   "Form2"
   ClientHeight    =   6705
   ClientLeft      =   -120
   ClientTop       =   2475
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6705
   ScaleWidth      =   11385
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Chart Styles"
      Height          =   4455
      Left            =   8880
      TabIndex        =   1
      Tag             =   "noprint"
      Top             =   600
      Width           =   2295
      Begin VB.OptionButton Option1 
         Caption         =   "2dBar"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   11
         Top             =   3720
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3dStep"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   10
         Top             =   3360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3dBar"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   9
         Top             =   3000
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3dLine"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   8
         Top             =   2640
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2dLine"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2dXY"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2dStep"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2dPie"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2dCombination"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3dArea"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
   End
   Begin MSChartLib.MSChart Chart1 
      DragMode        =   1  'Automatic
      Height          =   5895
      Left            =   120
      OleObjectBlob   =   "MsChartFrm.frx":0000
      TabIndex        =   0
      Top             =   600
      Width           =   8415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MsChart Styles"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "MsChartFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Option1_Click(Index As Integer)
 Chart1.RowCount = 4
    Chart1.ColumnCount = 5
Select Case Index
Case 0
 Chart1.chartType = VtChChartType2dCombination
 Case 1
  Chart1.chartType = VtChChartType3dArea
  Case 2
 Chart1.chartType = VtChChartType2dPie
 Case 3
  Chart1.chartType = VtChChartType2dStep
  Case 4
  Lulaa 'Chart1.chartType = VtChChartType2dXY
  Case 5
  Chart1.chartType = VtChChartType2dLine
  Case 6
  Chart1.chartType = VtChChartType3dLine
  Case 7
  Chart1.chartType = VtChChartType3dBar
  Case 8
  Chart1.chartType = VtChChartType3dStep
  Case 9
 Chart1.chartType = VtChChartType2dBar
 End Select
End Sub

Private Sub Lulaa()
Dim theta_min As Single
Dim theta_max As Single
Dim dtheta As Single
Dim theta As Single
Dim r As Single
Dim values() As Single
Dim i As Integer
Dim num_theta As Integer

    theta_min = 0
    theta_max = 3.14159265
    num_theta = 100
    dtheta = theta_max / (num_theta - 1)
    ReDim values(1 To num_theta, 1 To 2)

    ' Compute the data values.
    theta = theta_min
    For i = 1 To num_theta
        r = Cos(3 * theta)
        values(i, 1) = r * Sin(theta) * 100
        values(i, 2) = r * Cos(theta) * 100
        theta = theta + dtheta
    Next i

    ' Send the data to the chart.
    Chart1.chartType = VtChChartType2dXY
    Chart1.RowCount = 2
    Chart1.ColumnCount = num_theta
    Chart1.ChartData = values
End Sub
