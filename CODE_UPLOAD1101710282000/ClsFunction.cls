VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFunction"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public oConnection As New ADODB.Connection
Public rs As Recordset
Public db As Database
Public isfonction As Boolean
Public isfin As Boolean
Public X As Long
Public Y As Long

Public Function Eval(Champ As Variant, isfonction As Boolean) As Variant
Dim rs2 As Recordset
Dim i As Integer
Dim j As Integer
Dim flag As Boolean
Dim flag2 As Boolean
Dim VraiChamp As Variant
Dim a_calculer As String
Dim Result As Variant
Dim sFormat As Variant
Dim textetemp As Variant
Dim a As Integer
Dim Flag_Fnct As Boolean
Dim txtfnt As Variant
Dim compteur As Integer
Dim matrice() As Variant
Dim Fonction_Name As String
Dim sSql As String
Dim param() As Variant
Dim cParam As Integer
Dim txtparam As Variant
Dim fct As Variant
Dim tmp As Variant
Dim txttmp As Variant
Dim Emplir_Param As Boolean
Dim valeur As Variant
Dim txtfnct As Variant
Dim Champ_Param As Variant
Dim param1 As Variant
Dim param2 As Variant
Dim objet As Object
Dim temp As Variant
Dim texte_evaluer As Variant
Dim valeur_fonction As Variant
Dim matrice_fonction() As String
Dim matrice_param() As String
Dim Cptr As Long
Dim parametre() As String
Dim flag_passer As Boolean
Dim X As Long
Dim Y As Long
Dim temp2 As Variant
Dim oMyParser As New clsExpressionParser
Set oMyParser = New clsExpressionParser
Dim clsfonction As New ClsFunction
Set clsfonction = New ClsFunction
Set rs2 = db.OpenRecordset("larequete")

    
If Not isfonction Then
    For a = 1 To Len(Champ)
    If Not IsNull(Mid(Champ, a, 1)) Then
        Select Case Mid(Champ, a, 1)
            Case "[", "]"
                If Mid(Champ, a, 1) = "[" Then
                    If Not flag Then
                        flag = True
                    Else
                        flag = False
                    End If
                    Emplir_Param = True
                Else
                    If Not Flag_Fnct Then
                        If flag = True And Mid(Champ, a, 1) = "]" Then
                            VraiChamp = textetemp
                            valeur = rs2(Champ_Param) & ""
                            If valeur = "" Then
                                valeur = 0
                            End If
                            a_calculer = a_calculer & valeur
                            Champ_Param = ""
                            textetemp = ""
                            flag = False
                            Emplir_Param = False
                        Else
                            a_calculer = a_calculer & Mid(Champ, a, 1)
                        End If
                    Else
                        If Flag_Fnct Then
                            txtfnct = txtfnct & CStr(rs2(Champ_Param))
                        End If
                        valeur = rs2(Champ_Param)
                        rs2.Close
                        Champ_Param = ""
                        Emplir_Param = False
                    End If
                End If
            Case Chr(34)
                 If Flag_Fnct = True Then
                        txtfnct = txtfnct & Chr(34)
                Else
                    If flag = False Then
                        flag = True
                    Else
                       flag = False
                    End If
                End If
                
            Case Chr(123), Chr(125)
                Flag_Fnct = True
                fct = Eval_Fnc(Mid(Champ, a))
'                tmp = Eval(fct, isfonction)
                Champ = fct
                If InStr(Champ, "{") > 0 Then
                    Champ = Eval_Fnc(Champ)
                Else
                    Eval = Champ
                    Exit Function
                End If
                a = a + Len(fct)
                
                
                If Mid(Champ, a, 1) = Chr(123) Then
                    Flag_Fnct = True
                Else
                    If Flag_Fnct And Mid(Champ, a, 1) = Chr(125) Then
                        For compteur = 0 To UBound(param()) - 1
                            If compteur < UBound(param()) - 1 Then
                                txtparam = txtparam & param(compteur) & ", "
                            Else
                                txtparam = txtparam & param(compteur)
                            End If
                        Next compteur
                        matrice() = Split(txtfnct, "(")
                        Fonction_Name = matrice(0)
                        MsgBox Fonction_Name
                        MsgBox txtparam
'                        Flag_Fnct = False
                        param1 = param(0)
                        param2 = param(1)
                        temp = clsfonction.exec_fonction(objet, Fonction_Name, CStr(param1) & ", " & CStr(param2) & ", " & "allo")
                        'TEMP = CallByName(Me, Fonction_Name, 1, txtparam)
                    Else
                        txtfnct = txtfnct & Mid(Champ, a, 1)
                    End If
                End If
            Case "(", ")", "+", "-", "/", "*", Chr(48), Chr(49), Chr(50), Chr(51), Chr(52), Chr(53), Chr(54), Chr(55), Chr(56), Chr(57)
                If Not Flag_Fnct Then
                    
                    a_calculer = a_calculer & Mid(Champ, a, 1)
                Else
                    txtfnct = txtfnct & Mid(Champ, a, 1)
                End If
            Case Else
                If flag Or Flag_Fnct Or IsNumeric(Mid(Champ, a, 1)) Then
                    If Not flag Then
                        textetemp = textetemp & Mid(Champ, a, 1)
                        If flag Then
                            a_calculer = a_calculer & Mid(Champ, a, 1)
                        End If
                    Else
                        If Emplir_Param Then
                            Champ_Param = Champ_Param & Mid(Champ, a, 1)
                        Else
                            txtfnct = txtfnct & Mid(Champ, a, 1)
                        End If
                   
                    End If
                Else
                    a_calculer = a_calculer & Mid(Champ, a, 1)
                End If
               ' textetemp = Replace(textetemp, Chr(34), "")
                
            End Select
            
        End If
    Next a
Else
    On Error GoTo erreur_fonction
    If InStr(Champ, "(") > 0 Then
        matrice_fonction() = Split(Champ, "(")
        If Len(matrice_fonction(1)) Then
            matrice_param() = Split(matrice_fonction(1), ",")
            cParam = 0
            For Cptr = 0 To UBound(matrice_param())
                ReDim Preserve parametre(Cptr)
                If Cptr = 0 Then
                    parametre(Cptr) = format_parametre(matrice_param(Cptr))
                Else
                    parametre(Cptr) = format_parametre(matrice_param(Cptr))
                End If
            Next Cptr
            a_calculer = clsfonction.exec_fonction_matrice(clsfonction, matrice_fonction(0), parametre())
        Else
            MsgBox "Champs manquant dans la fonction"
        End If
    Else
        isfonction = False
    End If
End If

If Len(a_calculer) And Not flag_passer Then
    'MsgBox a_calculer
'    If InStr(a_calculer, "(") > 0 Then
'        isfonction = True
'        temp = Eval(a_calculer, isfonction)
'    End If
    If Val(a_calculer) > 0 Then
        Result = oMyParser.ParseExpression(a_calculer)
    Else
        Result = a_calculer
    End If
    'MsgBox "Résultat : " & Format(Result, sFormat)
    flag_passer = True
End If



Result = 0
VraiChamp = ""
'a_calculer = ""
textetemp = ""
If a_calculer <> "" Then
    If InStr(a_calculer, "+") > 0 Or InStr(a_calculer, "-") > 0 Or InStr(a_calculer, "/") > 0 Or InStr(a_calculer, "*") > 0 Then
        Call Me.Imprime(oMyParser.ParseExpression(a_calculer))
    Else
        Call Me.Imprime(a_calculer)
    End If
ElseIf Result <> "" And Result <> 0 Then
Call Me.Imprime(Result)
ElseIf txtfnct <> "" Then
    Call Me.Imprime(txtfnct)
End If
textetemp = ""
Champ = ""
Exit Function
erreur_fonction:
If Err.Number = 13 Then
    MsgBox "Erreur dans l'exécution de la fonction" & vbCrLf & "Erreur : Champ manquant"
Else
    MsgBox "Erreur dans l'exécution de la fonction" & vbCrLf & "Erreur : " & Err.Description
End If
End

End Function


Public Function Eval_Fnc(texte As Variant, Optional isfin As Boolean) As String

Dim Stack As Integer
Dim i As Long
Dim posidebut As Integer
Dim A_Verifier As String
Dim Flag_Fonction As Boolean
Dim EnCours As Integer
Dim posfin As Integer
Dim Emplir_Param As Boolean
Dim txtparam As String
Dim txtresult As Variant
Dim matrice() As String
Dim Fonction_Name  As String
Dim param1 As Variant
Dim param2 As Variant
Dim param() As String
Dim temp As Variant
Dim j As Long
Dim a As Integer
Dim X As Integer
Dim Y As Integer
Dim z As Integer
Dim var_extraire As String
Dim var_extrait As String
Dim matrice2() As String
Dim reste As String
Dim lavaleur As String
Dim fin As Boolean


For i = 1 To Len(texte)
    
    Select Case Mid(texte, i, 1)

    Case "{"
        'I = posfin
        A_Verifier = A_Verifier & Mid(texte, i, 1)
        Flag_Fonction = True
        Stack = Stack + 1
    Case "}"
        A_Verifier = A_Verifier & Mid(texte, i, 1)
        Stack = Stack - 1
        If Stack < 0 Then
            MsgBox "Erreur d'accolade: } sans {"
            Exit Function
        Else
            If Stack = 0 Then
                isfonction = True
                If Left(A_Verifier, 4) = "{iif" Then
                    Eval_Fnc = Eval(A_Verifier, isfonction)
                ElseIf Left(A_Verifier, 6) = "isnull" Then
                        If Get_isnull_value(A_Verifier) = "null" Then
                            MsgBox "La réposne est: null"
                        Else
                            MsgBox "La réposne est: not null"
                        End If
                Else
                    matrice2() = Split(A_Verifier, "{")
                    For X = UBound(matrice2()) To 1 Step -1
                        Do While Not fin Or Y <> Len(matrice2(X))
                        Y = Y + 1
                        'For Y = 1 To Len(matrice2(X))
                            If Mid(matrice2(X), Y, 1) = "}" Then
                                matrice() = Split(var_extraire, "(")
                                For z = 1 To UBound(matrice())
                                    reste = reste & "(" & matrice(z)
                                Next z
                                Fonction_Name = matrice(0)
                                var_extrait = exec_fonction(Me, Fonction_Name, format_parametre(reste))
                                lavaleur = Replace(A_Verifier, "{" & Fonction_Name & reste & "}", var_extrait)
                                If InStr(lavaleur, "{") > 0 Then
                                    Eval_Fnc = Eval_Fnc(lavaleur, isfin)
                                    If isfin Then
                                        Exit Function
                                    End If
                                    
                                Else
                                    Eval_Fnc = lavaleur
                                    fin = True
                                    isfin = True
                                    var_extraire = ""
                                    var_extrait = ""
                                    lavaleur = ""
                                End If
                                If isfin Then
                                    GoTo fin
                                End If
                                '= Me.exec_fonction(Me, matrice2(0), var_extraire)
                            Else
                                var_extraire = var_extraire & Mid(matrice2(UBound(matrice2())), Y, 1)
                            End If
                        Loop
                    Next X
'                    Eval_Fnc = Me.exec_fonction_matrice(Me, Fonction_Name, param())
                    
                    'TEMP = CallByName(Me, Fonction_Name, 1, txtparam)
                        
                End If
                

                
            End If
        End If
    Case Chr(91)
        Emplir_Param = True
    Case Chr(93)
        If Emplir_Param Then
            A_Verifier = A_Verifier & Eval_Valeur(txtparam, "AS_CONTRAT_VENTE")
            txtparam = ""
            Emplir_Param = False
        End If
    Case Else
        If Not Emplir_Param Then
            A_Verifier = A_Verifier & Mid(texte, i, 1)
        Else
            txtparam = txtparam & Mid(texte, i, 1)
        End If
    End Select
Next i
   
fin:
End Function
Public Function Get_iif_value(texte2 As Variant) As Variant
Dim i As Integer
Dim txt_a_convertir As Variant
Dim a_convertir As Boolean
Dim chpboolean As Boolean
Dim matrice_conversion() As String
Dim j As Integer
Dim convertir As String
Dim temp As Variant
Dim valeur As String
Dim operande As String
Dim clsparser As New clsExpressionParser
Dim z As Integer
Dim posidebut As Long
Dim posifin As Long
Dim a_remplacer As String
Dim txt_a_convertir_temp As String
Dim valeur_traite As Variant
Dim Expr1 As Variant
Dim Expr2 As Variant
Dim Memoire_Operande As Variant
Dim posi_vrai_debut As Integer
Dim Si_vrai As String
Dim Si_non As String
Dim condition() As String

Set clsparser = New clsExpressionParser

z = 0

For i = 1 To Len(texte2)

    If Mid(texte2, i, 1) = "(" Then
        posi_vrai_debut = posi_vrai_debut + 1
        posidebut = i + 1
        a_convertir = True
    Else
        If a_convertir Then
'            matrice_conversion() = Split(txt_a_convertir, Chr(32))
'            For j = 0 To UBound(matrice_conversion())
                Select Case Mid(texte2, i, 1)
                Case "<"
                    Select Case Mid(texte2, i + 1, 1)
                    Case "<"
                        MsgBox "Opérateur non reconnu : << "
                        End
                        
                    Case ">"
                        valeur = clsparser.ParseExpression(convertir)
                        a_remplacer = Mid(texte2, posidebut, Len(convertir) + j)
                        posifin = Len(convertir) + posidebut + 1
                        txt_a_convertir = Replace(convertir, a_remplacer, valeur)
                        i = posifin
                        posidebut = posifin + 1
                        Expr1 = valeur
                        Memoire_Operande = "<>"
                        z = z + 1
                        convertir = ""
                        'txt_a_convertir = ""
                    Case "="
                        valeur = clsparser.ParseExpression(convertir)
                        a_remplacer = Mid(texte2, posidebut, Len(convertir) + j)
                        posifin = Len(convertir) + posidebut + 1
                        txt_a_convertir = Replace(convertir, a_remplacer, valeur)
                        i = posifin
                        posidebut = posifin + 1
                        Expr1 = valeur
                        Memoire_Operande = "<="
                        z = z + 1
                        convertir = ""
                        'txt_a_convertir = ""
                    Case Chr(32)
                        valeur = clsparser.ParseExpression(convertir)
                        a_remplacer = Mid(texte2, posidebut, Len(convertir) + j)
                        posifin = Len(convertir) + posidebut + 1
                        txt_a_convertir = Replace(convertir, a_remplacer, valeur)
                        i = posifin
                        posidebut = posifin + 1
                        Expr1 = valeur
                        Memoire_Operande = "<"
                        z = z + 1
                        convertir = ""
                        'txt_a_convertir = ""
                    Case Else
                    End Select
                Case ">"
                    Select Case Mid(texte2, i + 1, 1)
                    Case "<"
                        MsgBox "Opérateur non reconnu : >< "
                        End
                        
                    Case ">"
                        MsgBox "Opérateur non reconnu : >> "
                        End
                    Case "="
                        valeur = clsparser.ParseExpression(convertir)
                        a_remplacer = Mid(texte2, posidebut, Len(convertir) + j)
                        posifin = Len(convertir) + posidebut + 1
                        txt_a_convertir = Replace(convertir, a_remplacer, valeur)
                        i = posifin
                        posidebut = posifin + 1
                        Expr1 = valeur
                        Memoire_Operande = ">="
                        z = z + 1
                        convertir = ""
                        'txt_a_convertir = ""
                    Case Chr(32)
                        valeur = clsparser.ParseExpression(convertir)
                        a_remplacer = Mid(texte2, posidebut, Len(convertir) + j)
                        posifin = Len(convertir) + posidebut + 1
                        txt_a_convertir = Replace(convertir, a_remplacer, valeur)
                        i = posifin
                        posidebut = posifin + 1
                        Expr1 = valeur
                        Memoire_Operande = ">"
                        z = z + 1
                        convertir = ""
                    Case Else
                    End Select
                   
                Case "="
                    convertir = convertir & Mid(texte2, i, 1)
                    Select Case Mid(texte2, i + 1, 1)
                        Case "<"
                            MsgBox "Opérateur non reconnu : << "
                            End

                        Case ">"
                            MsgBox "Opérateur non reconnu : => "
                            End
                        Case "="
                            MsgBox "Opérateur non reconnu: == "
                        Case Chr(32)
                            valeur = clsparser.ParseExpression(convertir)
                            a_remplacer = Mid(texte2, posidebut, Len(convertir) + j)
                            posifin = Len(convertir) + posidebut
                            txt_a_convertir = Replace(convertir, a_remplacer, valeur)
                            i = posifin
                            posidebut = posifin + 1
                            Expr1 = valeur
                            Memoire_Operande = "="
                            z = z + 1
                            convertir = ""
                        Case Else
                    End Select
                Case ","
                    valeur = clsparser.ParseExpression(convertir)
                    a_remplacer = Mid(texte2, posidebut, Len(convertir) + j)
                    posifin = Len(convertir) + posidebut
                    txt_a_convertir = txt_a_convertir & Memoire_Operande & Replace(convertir, a_remplacer, valeur)
                    i = posifin
                    posidebut = posifin
                    condition() = Split(Mid(texte2, posifin), ",")
                    Si_vrai = condition(1)
                    Si_non = Replace(condition(2), ")", "")
                    Expr2 = valeur
                    convertir = ""
                    'txt_a_convertir = ""
                    Get_iif_value = eval_expres_iif(Expr1, Memoire_Operande, Expr2, Si_vrai, Si_non)
                    texte2 = Replace(texte2, Mid(texte2, posi_vrai_debut, posifin - posi_vrai_debut), Get_iif_value)
                    texte2 = Replace(texte2, ")", "")
                    Exit Function
                Case Else
                    convertir = convertir & Mid(texte2, i, 1)
                End Select
        'Next j
'        Else
'            If a_convertir Then
'                txt_a_convertir = txt_a_convertir & Mid(texte2, i, 1)
'            End If
        Else
            
        End If
        
    End If
Next i
If InStr(convertir, "iif") > 0 Then
    temp = Split_Param(convertir, ",", condition())
End If
End Function

Public Function calcul(chp1 As Integer, chp2 As Integer) As Integer
Dim temp As Integer
calcul = chp1 + chp2

End Function
Public Function OuvrirBd(labd As Database, lers As Recordset, Nom_base As String, Nom_recordset As String) As Boolean

On Error GoTo erreur
Set db = OpenDatabase(Nom_base)
Set lers = db.OpenRecordset(Nom_recordset)

Set rs = lers
Set labd = db
OuvrirBd = True
Exit Function
erreur:
OuvrirBd = False

End Function
Public Function format_parametre(Leparametre As Variant) As Variant
Dim txttmp As Variant
Dim txttmp2 As Variant

txttmp = Replace(Leparametre, Chr(34), "")
txttmp = Replace(txttmp, "(", "")
txttmp = Replace(txttmp, ")", "")
txttmp = Replace(txttmp, "{", "")
txttmp = Replace(txttmp, "}", "")
If Not IsNull(txttmp) Then
    format_parametre = txttmp
Else
    format_parametre = Chr(34) & Chr(34)
End If

End Function

Public Function eval_expres_iif(expression1 As Variant, operande As Variant, expression2 As Variant, condition_vrai As String, condition_fausse As String) As Variant

Select Case operande
    Case "<"
        If expression1 < expression2 Then
            eval_expres_iif = condition_vrai
        Else
            eval_expres_iif = condition_fausse
        End If
    Case "<="
        If expression1 <= expression2 Then
            eval_expres_iif = condition_vrai
        Else
            eval_expres_iif = condition_fausse
        End If
    Case ">"
    If expression1 > expression2 Then
            eval_expres_iif = condition_vrai
        Else
            eval_expres_iif = condition_fausse
        End If
    Case ">="
        If expression1 >= expression2 Then
            eval_expres_iif = condition_vrai
        Else
            eval_expres_iif = condition_fausse
        End If
    Case "="
        If expression1 = expression2 Then
            eval_expres_iif = condition_vrai
        Else
            eval_expres_iif = condition_fausse
        End If
    Case "<>"
        If expression1 <> expression2 Then
            eval_expres_iif = condition_vrai
        Else
            eval_expres_iif = condition_fausse
        End If
    End Select
    
End Function

Public Function valide_structure(a_valider As Variant) As Boolean
Dim parenthese As Integer
Dim virgule As Integer
Dim braquette As Integer
Dim i As Integer
For i = 1 To Len(a_valider)
    Select Case Mid(a_valider, i, 1)
    Case "("
        parenthese = parenthese + 1
    Case ")"
        parenthese = parenthese - 1
    Case ","
        virgule = virgule + 1
    Case "["
        braquette = braquette + 1
    Case "]"
        braquette = braquette - 1
    End Select
Next i

If parenthese <> 0 Or braquette <> 0 Or virgule < 2 Then
    valide_structure = False
Else
    valide_structure = True
End If

End Function
Public Function Get_isnull_value(valeur As Variant) As Variant
Dim clsparser As New clsExpressionParser
Dim i As Integer
Dim posidebut As Integer
Dim posifin As Integer
Set clsparser = New clsExpressionParser
posifin = 1
posidebut = 8


For i = 1 To Len(valeur)
    If Mid(valeur, i, 1) = ")" Then
        posifin = i - 1
    End If
Next i
If posifin > posidebut Then
    If Not IsNull(clsparser.ParseExpression(Mid(valeur, posidebut, posifin - posidebut))) Then
        Get_isnull_value = "non null"
    End If
Else
    Get_isnull_value = "null"
End If
End Function



Public Function new_use_test(Type_Vehicule As Long, Field1 As String) As String

'Field1 = Nom du field de la case NEUVE

Select Case Field1
Case "Neuve"
    If Type_Vehicule = 0 Then   'type de vehicule
        Field1 = "X"
    Else
        If Type_Vehicule = 1 Then
            Field1 = ""
        Else
            Field1 = ""
        End If
    End If
Case "Usagee"
    If Type_Vehicule = 0 Then   'type de vehicule
        Field1 = ""
    Else
        If Type_Vehicule = 1 Then
            Field1 = "X"
        Else
            Field1 = ""
        End If
    End If
Case "Essai"
    If Type_Vehicule = 0 Then   'type de vehicule
        Field1 = ""
    Else
        If Type_Vehicule = 1 Then
            Field1 = ""
        Else
            Field1 = "X"
        End If
    End If
End Select

End Function
'Public Function coche_si_radio() As String
'If TrouverProduit([idintveh], 473) > -1 Or TrouverProduit([idintveh], 474) > -1 Or TrouverProduit([idintveh], 475) > -1 Or TrouverProduit([idintveh], 503) > -1 Then
'    coche_si_radio = "X"
'End If
'End Function
Public Function TrouverProduit_montant(ID_vehicule As Variant, Piece As Variant, Cat_Field As Variant) As Variant
Dim matrice() As String

Select Case Trim(Cat_Field)
Case "X"
    If TrouverProduit(Val(ID_vehicule), Val(Piece)) > -1 Then
        TrouverProduit_montant = "X"
    End If
Case "Montant"
        
    If TrouverProduit(Val(ID_vehicule), Val(Piece)) > -1 Then
        
        matrice() = Split(TrouverProduit(Val(ID_vehicule), Val(Piece)), "þ")
        If Cat_Field > 0 Then
            TrouverProduit_montant = Format(matrice(0), "#,##0.00")
        End If
    End If
Case "Nom"
    If TrouverProduit(Val(ID_vehicule), Val(Piece)) > -1 Then
        If InStr(TrouverProduit(Val(ID_vehicule), Val(Piece)), "þ") > 0 Then
            matrice() = Split(TrouverProduit(Val(ID_vehicule), Val(Piece)), "þ")
            If Cat_Field > 0 Then
                TrouverProduit_montant = matrice(1)
            End If
        Else
                TrouverProduit_montant = ""
        End If
    End If
End Select


End Function
Public Function TrouverProduit(ID_vehicule As String, ID_Produit As String) As Variant

Dim rstContrat As ADODB.Recordset
Dim rstTemp1 As ADODB.Recordset
Dim sSql9 As String
'Dim oConnection As New ADODB.Connection
If oConnection.State <> adStateOpen Then
    oConnection.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Autosoft"
End If
Set rstTemp1 = New ADODB.Recordset
    Set rstContrat = New ADODB.Recordset

    Dim oCmd As ADODB.Command
    Dim oParam1 As ADODB.Parameter
    
    'initialisation des paramètres
    Set oCmd = New ADODB.Command
    Set oParam1 = New ADODB.Parameter
    oParam1.Type = adInteger
    oParam1.Direction = adParamInput
    'oParam1.Value = ID_Contrat

    'ouverture de la vue logique
    oCmd.Parameters.Append oParam1
    oCmd.CommandText = "AS_CONTRAT_VENTE"
    Set oCmd.ActiveConnection = oConnection
    'Set rstContrat = oCmd.Execute
    
sSql9 = "SELECT ASINTVEHPRO.IDINTVEH, ASINTVEHPRO.IDPRO, ASINTVEHPRO.MONTANT, GRPRO.DESC0" & _
        " FROM ASINTVEHPRO INNER JOIN GRPRO ON ASINTVEHPRO.IDPRO = GRPRO.ID" & _
        " WHERE (((ASINTVEHPRO.IDINTVEH)=175) AND ((ASINTVEHPRO.IDPRO)=123));"

rstTemp1.Open sSql9, oConnection

If Not (rstTemp1.EOF) Then
    TrouverProduit = rstTemp1!Montant & "þ" & rstTemp1!desc0
Else
    TrouverProduit = -1
End If

'If rstTemp1.State = adStateOpen Then rstTemp1.Close
Set rstTemp1 = Nothing

End Function
Public Function Concatener(ID_vehicule As Long) As String
'Retourne sous forme de chaine la liste des produits reliés à un véhicule.
'Si le véhicule en question n'a pas de produits sélectionnées la chaine
'   "Pas d'options." est retournée.

'Entrée : L'id du véhicule
'Sortie : Liste des produits

    Dim strResultat
    Dim rstTemp As ADODB.Recordset
    Dim sSql As String
    
    'préparation de la requete qui va chercher toutes les options rattachées au vehicule
    Set rstTemp = New ADODB.Recordset
    sSql = "SELECT ASINTVEHPRO.IDINTVEH, sysUSRCOD.DESC0 " & _
    "FROM ASINTVEHPRO INNER JOIN sysUSRCOD ON ASINTVEHPRO.IDPRO = sysUSRCOD.ID " & _
    "WHERE (((ASINTVEHPRO.IDINTVEH)= " & ID_vehicule & "));"
    
    rstTemp.Open sSql, oConnection
    
    Do Until rstTemp.EOF
        strResultat = strResultat & rstTemp.Fields(1)
        rstTemp.MoveNext
        If rstTemp.EOF Then
            strResultat = strResultat & "."
        Else
            strResultat = strResultat & ", "
        End If
    Loop
    
    If Len(strResultat) = 0 Then
        Concatener = "Pas d'options."
    Else
        Concatener = strResultat
    End If
    
    If rstTemp.State = adStateOpen Then rstTemp.Close
    Set rstTemp = Nothing
    
End Function
Public Function TrouverAccessoires(ID_vehicule As Long, strListe As String, intMontantTotal As Long)
' Ne doit plus etre utlisé     À VÉRIFIER

    Dim strYO As String
    Dim rstTemp1 As ADODB.Recordset
    Dim sSql10 As String
    
    strYO = "Complémentaire"
    'préparation de la requete qui va chercher toutes les options rattachées au vehicule
    Set rstTemp1 = New ADODB.Recordset
    sSql10 = "SELECT ASINTVEHPRO.IDINTVEH, sysUSRCOD.VALEURCAR, ASINTVEHPRO.IDPRO, ASINTVEHPRO.MONTANT, sysUSRCOD.DESC0 " & _
            "FROM ASINTVEHPRO INNER JOIN sysUSRCOD ON ASINTVEHPRO.IDPRO = sysUSRCOD.ID " & _
            "WHERE (((ASINTVEHPRO.IDINTVEH)= " & ID_vehicule & " ) AND ((sysUSRCOD.VALEURCAR)= '" & strYO & "'));"
    
    rstTemp1.Open sSql10
    ', oConnection
    
    Do Until rstTemp1.EOF
        strListe = strListe & rstTemp1!desc0
        intMontantTotal = intMontantTotal + rstTemp1!Montant
        rstTemp1.MoveNext
        If rstTemp1.EOF Then
            strListe = strListe & "."
        Else
            strListe = strListe & ", "
        End If
    Loop
    
    If rstTemp1.State = adStateOpen Then rstTemp1.Close
    Set rstTemp1 = Nothing

End Function

Public Function TrouverConcessionnaire(strNom As String, strRue As String, strVille As String, strProv As String, strCodePostal As String, strTPS As String, strTVQ As String)
'Cette fonction retourne les informations concenant ler concessionnaire.

'Entrée : aucune (L'enregistrement 1 de sysCIE DOIT représenter le concessionnaire en question)
'Sortie : Le nom (strNom), l'adresse (strRue), la ville(strVille),  la province(strProv),
'       le code postal(strCodePostal), le noméro de tps(strTPS) et le numéro de tvq(strTVQ)
    
    Dim rstTemp2 As ADODB.Recordset
    Dim sSql5 As String
    
    Set rstTemp2 = New ADODB.Recordset
    sSql5 = "SELECT sysCIE.ID, sysCIE.NOM0, sysCIE.NORUE, sysCIE.RUE, sysCIE.SUITE, sysCIE.VILLE, sysCIE.PAYS, sysCIE.PROVETAT, sysCIE.CODEPOSTAL, sysCIE.NOTEL, sysCIE.NOTAXE1, sysCIE.NOTAXE2 " & _
            "From sysCIE " & _
            "WHERE (((sysCIE.ID)=1));"
            
    rstTemp2.Open sSql5
    ', oConnection

    strNom = rstTemp2!nom0
    strRue = rstTemp2!norue & rstTemp2!rue
    strVille = rstTemp2!ville
    strProv = rstTemp2!provetat
    strCodePostal = rstTemp2!codepostal
    strTPS = rstTemp2!notaxe1
    strTVQ = rstTemp2!notaxe2

    If rstTemp2.State = adStateOpen Then rstTemp2.Close
    Set rstTemp2 = Nothing

End Function

Public Function ExtraireTaxes(Montant As Double, TauxTPS As Double, TauxTVQ As Double, NoTax As Double, TPS As Double, TVQ As Double)
'Cette fonction retourne les frais de tps (TPS), de tvq (TVQ) ainsi que le montant
'avant taxe (NoTax) d'un montant qui inclu les taxes (Montant).

'Entrée : Montant qui inclue les taxes (Montant) Taux de la TPS (TauxTPS) Taux de TVQ (TauxTVQ)
'Sortie : Montant sans taxe (NoTax) Montant de TPS (TPS) Montant de TVQ (TVQ)

    Dim TaxeTotal As Double
    Dim temp1 As Double
    Dim temp2 As Double

    TaxeTotal = (1 + (Int(TauxTVQ * 10 + 0.5) / 10 / 100)) * (1 + (Int(TauxTPS * 10 + 0.5) / 10 / 100))
    NoTax = Int(Montant / TaxeTotal * 100 + 0.5) / 100
    temp1 = NoTax * TauxTPS / 100
    TPS = Int(temp1 * 100 + 0.5) / 100
    temp2 = NoTax * TauxTVQ / 100
    TVQ = Int(temp2 * 100 + 0.5) / 100

End Function

Public Function chiffre(texte As Variant) As Boolean
If Asc(texte) < 47 Or Asc(texte) > 57 Then
    chiffre = False
Else
    chiffre = True
End If
End Function

Public Function garantie(Annee As Long, Km As Long, Type_veh As Long, Lettre_case As String) As Boolean

Dim conn As ADODB.Connection
Dim rsconn As ADODB.Recordset
Dim sSql As String
Dim flag As Boolean

sSql = "SELECT codelpc from AS_GARANTIE_USAGE_LPC where " & _
"AS_GARANTIE_USAGE_LPC.bornekmmin <= " & Km & " and " & _
"AS_GARANTIE_USAGE_LPC.borneagevehmax <= " & Annee & "and " & _
" AS_GARANTIE_USAGE_LPC.tpvehicule = " & Type_veh & ";"

Set conn = New ADODB.Connection
conn.Open ("Autosoft")
Set rsconn = conn.Execute(sSql)
Do Until rsconn.EOF
    If UCase(rsconn.Fields(0)) = UCase(Lettre_case) Then
        flag = True
    End If
rsconn.MoveNext
Loop
rsconn.Close
conn.Close

garantie = flag

End Function


Public Function exec_fonction(Object_Name As Project1.ClsFunction, Fonction_Name As String, texte As String) As Variant
Dim i As Integer
Dim matrice() As String
On Error Resume Next
matrice() = Split(texte, ",")

'If Len(matrice()) Then

Select Case UBound(matrice())
    Case 0
        exec_fonction = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0))
    Case 1
        exec_fonction = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1))
    Case 2
        exec_fonction = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2))
    Case 3
        exec_fonction = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2), matrice(3))
    Case 4
        exec_fonction = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2), matrice(3), matrice(4))
    Case 5
        exec_fonction = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2), matrice(3), matrice(4), matrice(5))
    Case 6
        exec_fonction = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2), matrice(3), matrice(4), matrice(5), matrice(6))
    Case 7
        exec_fonction = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2), matrice(3), matrice(4), matrice(5), matrice(6), matrice(7))
    Case 8
        exec_fonction = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2), matrice(3), matrice(4), matrice(5), matrice(6), matrice(7), matrice(8))
    Case 9
        exec_fonction = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2), matrice(3), matrice(4), matrice(5), matrice(6), matrice(7), matrice(8), matrice(9))
    Case Else
    MsgBox "Nombre inexacte de paramètre ( 1 à 10 )"
End Select

End Function
Public Function exec_fonction_matrice(Object_Name As Object, Fonction_Name As String, matrice() As String) As Variant
Dim i As Integer
Dim temp As Variant
Dim expression1 As Variant
Dim expression2 As Variant
Dim operande As Variant
Dim split_expression() As String


'If Len(matrice()) Then
Fonction_Name = Replace(Fonction_Name, "{", "")
If Fonction_Name = "iif" Then
    Fonction_Name = "eval_expres_iif"
    split_expression() = Split(matrice(0), "=")
    expression1 = split_expression(0)
    operande = "="
    expression2 = split_expression(1)


    Select Case UBound(matrice())
        Case 0
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0))
        Case 1
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, CStr(matrice(0)), CStr(matrice(1)))
        Case 2
            exec_fonction_matrice = CallByName(Me, Fonction_Name, VbMethod, expression1, operande, expression2, matrice(1), matrice(2))
        Case 3
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, expression1, operande, expression2, matrice(1), matrice(2), matrice(3))
        Case 4
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, expression1, operande, expression2, matrice(1), matrice(2), matrice(3), matrice(4))
        Case 5
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, expression1, operande, expression2, matrice(1), matrice(2), matrice(3), matrice(4), matrice(5))
        Case 6
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, expression1, operande, expression2, matrice(1), matrice(2), matrice(3), matrice(4), matrice(5), matrice(6))
        Case 7
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, expression1, operande, expression2, matrice(1), matrice(2), matrice(3), matrice(4), matrice(5), matrice(6), matrice(7))
        Case 8
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, expression1, operande, expression2, matrice(1), matrice(2), matrice(3), matrice(4), matrice(5), matrice(6), matrice(7), matrice(8))
        Case 9
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, expression1, operande, expression2, matrice(1), matrice(2), matrice(3), matrice(4), matrice(5), matrice(6), matrice(7), matrice(8), matrice(9))
        Case Else
        MsgBox "Nombre inexacte de paramètre ( 1 à 10 )"
    End Select

Else
    Select Case UBound(matrice())
        Case 0
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0))
        Case 1
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1))
        Case 2
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2))
        Case 3
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2), matrice(3))
        Case 4
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2), matrice(3), matrice(4))
        Case 5
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2), matrice(3), matrice(4), matrice(5))
        Case 6
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2), matrice(3), matrice(4), matrice(5), matrice(6))
        Case 7
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2), matrice(3), matrice(4), matrice(5), matrice(6), matrice(7))
        Case 8
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2), matrice(3), matrice(4), matrice(5), matrice(6), matrice(7), matrice(8))
        Case 9
            exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0), matrice(1), matrice(2), matrice(3), matrice(4), matrice(5), matrice(6), matrice(7), matrice(8), matrice(9))
        Case Else
        MsgBox "Nombre inexacte de paramètre ( 1 à 10 )"
    End Select

End If
'MsgBox exec_fonction_matrice
End Function

Public Function unique_fonction(texte As String) As Boolean
Dim i As Long
Dim accolade As Integer

For i = 1 To Len(texte)
    If Mid(texte, i, 1) = "{" Then
        accolade = accolade + 1
    End If
Next i

If accolade > 0 Then
    unique_fonction = False
Else
    unique_fonction = True
End If

End Function

Public Function Split_Param(texte As String, Delemiteur As String, param() As String) As Variant
Dim i As Long
Dim j As Integer
Dim Resultat As String
Dim a_extraire As String
Dim accolade As Integer

For i = 1 To Len(texte)
    Select Case Mid(texte, i, 1)
    Case "{"
        a_extraire = a_extraire & Mid(texte, i, 1)
        accolade = accolade + 1
    Case "}"
        a_extraire = a_extraire & Mid(texte, i, 1)
        accolade = accolade - 1
        If accolade = 0 Then
            Resultat = Eval_Fnc(a_extraire)
            a_extraire = ""
        End If
    Case Else
        a_extraire = a_extraire & Mid(texte, i, 1)
    End Select
Next i
If Len(a_extraire) Then
    Resultat = Resultat & a_extraire
End If
param() = Split(Resultat, Delemiteur)
End Function
Public Function calc(var1 As String, var2 As String) As Integer
calc = Val(var1) + Val(var2)
End Function

Public Function sous(var1 As String, var2 As String) As Integer
sous = Val(var1) + Val(var2)
End Function

Public Sub Imprime(Letexte As Variant)
If Letexte <> "" And Letexte <> "0" Then
    Printer.CurrentX = Module1.X
    Printer.CurrentY = Module1.Y
    Printer.FontSize = Module1.Grosseur
    Printer.Print Letexte
End If
End Sub

