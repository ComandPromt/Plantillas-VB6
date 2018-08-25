VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Export To Excel"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   405
      Left            =   180
      TabIndex        =   1
      Top             =   2760
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid flxData 
      Height          =   2775
      Left            =   1440
      TabIndex        =   0
      Top             =   390
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   10
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oEXL              As Excel.Application


Private Sub cmdExport_Click()
    Export
End Sub

Private Sub Form_Load()
    
   flxData.TextMatrix(0, 0) = "Name"
   flxData.TextMatrix(0, 1) = "Age"
   flxData.TextMatrix(0, 2) = "ID"
   
   flxData.TextMatrix(1, 0) = "Guy"
   flxData.TextMatrix(1, 1) = "22"
   flxData.TextMatrix(1, 2) = "0378998"

   flxData.TextMatrix(2, 0) = "Ben"
   flxData.TextMatrix(2, 1) = "23"
   flxData.TextMatrix(2, 2) = "0123498"
   
   flxData.TextMatrix(3, 0) = "Gil"
   flxData.TextMatrix(3, 1) = "21"
   flxData.TextMatrix(3, 2) = "0325698"
   
   flxData.TextMatrix(4, 0) = "Shon"
   flxData.TextMatrix(4, 1) = "25"
   flxData.TextMatrix(4, 2) = "0325698"
    
   flxData.TextMatrix(5, 0) = "Geri"
   flxData.TextMatrix(5, 1) = "28"
   flxData.TextMatrix(5, 2) = "0563298"
    
   flxData.TextMatrix(6, 0) = "Mark"
   flxData.TextMatrix(6, 1) = "31"
   flxData.TextMatrix(6, 2) = "0563898"
    
   flxData.TextMatrix(7, 0) = "Dan"
   flxData.TextMatrix(7, 1) = "41"
   flxData.TextMatrix(7, 2) = "8756398"
    
   flxData.TextMatrix(8, 0) = "Kim"
   flxData.TextMatrix(8, 1) = "62"
   flxData.TextMatrix(8, 2) = "0325698"
    
   flxData.TextMatrix(9, 0) = "Nomi"
   flxData.TextMatrix(9, 1) = "31"
   flxData.TextMatrix(9, 2) = "04832198"
    
End Sub

Private Sub Export()
    
    Dim I           As Long
    Dim T           As Long
    
    
    Set oEXL = New Excel.Application
    With oEXL
        
        .Visible = True
        .Workbooks.Open App.Path & "\Export.xls"
        For I = 0 To Me.flxData.Rows - 1
            For T = 0 To (Me.flxData.Cols - 1)
                .Cells(I + 2, T + 1) = Me.flxData.TextMatrix(I, T)
            Next T
        
        Next I
        
    End With
        
        
        
        

End Sub






