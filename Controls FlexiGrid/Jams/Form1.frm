VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFG 
      Height          =   5445
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   9604
      _Version        =   393216
      Cols            =   4
      FillStyle       =   1
      SelectionMode   =   1
      FormatString    =   "^ |Description |>Date |>Destination"
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim i As Integer, tot As Integer
   Dim t As String, s As String

   ' Create sample data.
   t = Chr(9)
   MSFG.Rows = 1

   MSFG.AddItem "-" + t + "Airfare"
   s = "" + t + "SFO-JFK" + t + "9-Apr-95" + t + "750.00"
   For i = 0 To 5
      MSFG.AddItem s
   Next

   MSFG.AddItem "-" + t + "Meals"
   s = "" + t + "Flint's BBQ" + t + "25-Apr-95" + t + "35.00"
   For i = 0 To 5
      MSFG.AddItem s
   Next

   MSFG.AddItem "-" + t + "Hotel"
   s = "" + t + "Center Plaza" + t + "25-Apr-95" + t + "817.00"
   For i = 0 To 5
      MSFG.AddItem s
   Next

   ' Add up totals and format heading entries.
   For i = MSFG.Rows - 1 To 0 Step -1
      If MSFG.TextArray(i * MSFG.Cols) = "" Then
         tot = tot + Val(MSFG.TextArray(i * MSFG.Cols + 3))
      Else
         MSFG.Row = i
         MSFG.Col = 0
         MSFG.ColSel = MSFG.Cols - 1
         MSFG.CellBackColor = &HC0C0C0
         MSFG.CellFontBold = True
         MSFG.CellFontWidth = 8
         MSFG.TextArray(i * MSFG.Cols + 3) = _
         Format(tot, "0")
         tot = 0
      End If
   Next
   MSFG.ColSel = MSFG.Cols - 1

   ' Format Grid
   MSFG.ColWidth(0) = 300
   MSFG.ColWidth(1) = 1500
   MSFG.ColWidth(2) = 1000
   MSFG.ColWidth(3) = 1000

End Sub


Private Sub MSFG_DblClick()
   Dim i As Integer, r As Integer

   ' Ignore top row.
   r = MSFG.MouseRow
   If r < 1 Then Exit Sub

   ' Find field to collapse or expand.
   While r > 0 And MSFG.TextArray(r * MSFG.Cols) = ""
      r = r - 1
   Wend

   ' Show collapsed/expanded symbol on first column.
   If MSFG.TextArray(r * MSFG.Cols) = "+" Then
      MSFG.TextArray(r * MSFG.Cols) = "-"
   Else
      MSFG.TextArray(r * MSFG.Cols) = "+"
   End If

   ' Expand items under current heading.
   r = r + 1
   If MSFG.RowHeight(r) = 0 Then
      Do While MSFG.TextArray(r * MSFG.Cols) = ""
         MSFG.RowHeight(r) = -1 ' Default row height.
         r = r + 1
         If r >= MSFG.Rows Then Exit Do
      Loop

   ' Collapse items under current heading.
   Else
      Do While MSFG.TextArray(r * MSFG.Cols) = ""
         MSFG.RowHeight(r) = 0   ' Hide row.
         r = r + 1
         If r >= MSFG.Rows Then Exit Do
      Loop
   End If

End Sub
