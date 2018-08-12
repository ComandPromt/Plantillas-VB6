VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   3810
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oMyParser As New clsExpressionParser

' Keep track of user-defined constants
Dim ConstNames As New Collection
Public Function appel_fonction1(var1 As Variant, var2 As Variant) As Long
appel_fonction1 = var1 + var2
End Function

Private Sub Command1_Click()
'On Error Resume Next

Dim Emplir_Param As Boolean
Dim champ_fonction As Variant
Dim db As Database
Dim rs As Recordset
Dim rs2 As Recordset
Dim rs3 As Recordset
Dim clsfonction As ClsFunction
Dim param() As Variant
Dim cParam As Integer
Dim txtparam As String
Dim objet As Project1.clsExpressionParser
Dim Champ As Variant
Dim a_imprimer As Variant
Dim strsql As String
Initialise

strsql = "select * from latable where champ = " & InputBox("entrez l'id", "Le ID") & " order by ctltop"

Set clsfonction = New Project1.ClsFunction
Set objet = New Project1.clsExpressionParser
If clsfonction.OuvrirBd(db, rs, "lelecteur\labasededonnée", strsql) Then
    Grosseur = rs("ctlfontsize")
    clsfonction.Imprime (rs("IdSEC"))
    cParam = 0
    ReDim Preserve param(cParam)
    Do Until rs.EOF
        DoEvents
        
        X = rs("ctlleft")
        Y = rs("Ctltop")
        
        If Not IsNull(rs("ctldatafield")) Then
            'MsgBox rs("ctltag")
            clsfonction.isfonction = False
            Champ = rs("ctldatafield")
            If Not IsNull(rs4(Champ)) Then
                If Len(rs("ctloutputformat")) Then
                    clsfonction.Imprime Format(rs4(Champ), rs("ctloutputformat"))
                Else
                    clsfonction.Imprime rs4(Champ)
                End If
            End If
            'MsgBox " résultat : " & champ
        ElseIf Not IsNull(rs("ctltag")) Then
            'MsgBox rs("ctltag")
            
            clsfonction.isfonction = False
            Champ = clsfonction.Eval(rs("lechamptag"), clsfonction.isfonction)
            
            If Not IsNull(rs("ctloutputformat")) Then
                Champ = Format(Champ, rs("ctloutputformat"))
            End If
            'MsgBox " résultat : " & champ
        End If
        
    rs.MoveNext
    Loop
    rs.Close
    db.Close
    'Call clsfonction.Go_Imprime
    Printer.EndDoc
    Printer.PaperSize = 1
    MsgBox "tous les champs ont été parcourus"
    Exit Sub
Else
    MsgBox "erreur d'ouverture de la database"
    Exit Sub
End If
cmdParse_ErrHandler:
    
    If Err.Number >= PERR_FIRST And _
       Err.Number <= PERR_LAST Then
        ShowParseError
    Else
        MsgBox Err.Description, vbCritical, "Unexpected Error"
    End If

End Sub


Private Sub ShowParseError()
    
    ' Show error details
    MsgBox "Error No. " & CStr(Err.Number - PERR_FIRST + 1) & _
        " - " & Err.Description & vbCrLf & _
        "Raised from: " & Err.Source, _
        vbCritical, "Parse Error"
    
    ' Mark the position in the expression where the error
    ' was raised
    txtExpression.SelStart = oMyParser.LastErrorPosition - 1
    txtExpression.SelLength = 1
    txtExpression.SetFocus

End Sub
Private Sub Command2_Click()
Dim garantie As Variant

If garantie(2, 60000, 1, "c") Then
    MsgBox "X"
Else
    MsgBox " vide "
End If

End Sub

Private Sub Form_Load()
Printer.ScaleMode = 6 ' Millimètre
'Printer.ScaleTop = 1
'Printer.ScaleHeight = 135000
'Printer.PaperSize = 1
'Form2.Show

End Sub
