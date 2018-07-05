VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmData 
   Caption         =   "Data Designer Demo"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmData.frx":0000
      Height          =   4680
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   8255
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      AllowUserResizing=   3
      DataMember      =   "Orders"
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0)._NumMapCols=   5
      _Band(0)._MapCol(0)._Name=   "ShipCountry"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "CompanyName"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "OrderID"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(2)._Alignment=   7
      _Band(0)._MapCol(3)._Name=   "OrderDate"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "RequiredDate"
      _Band(0)._MapCol(4)._RSIndex=   4
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
    Dim i As Integer

    With MSHFlexGrid1

        .Redraw = False
'        ' set grid's column widths (-1 = default width)
        .ColWidth(0) = -1
        .ColWidth(1) = 2800         'company name
        .ColWidth(2) = -1
        .ColWidth(3) = -1
        .ColWidth(4) = -1

        ' set grid's column merging and sorting
        .MergeCells = flexMergeRestrictColumns
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next i

        .Sort = flexSortGenericAscending

        ' set grid's style
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat

        ' make header line with field names bold
        .Row = 0
        .Col = 0
        .RowSel = .FixedRows - 1
        .ColSel = .Cols - 1
        .CellFontBold = True

        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        .Redraw = True

    End With

End Sub
