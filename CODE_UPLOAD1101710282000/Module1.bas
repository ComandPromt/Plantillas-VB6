Attribute VB_Name = "Module1"



Public Function new_use_test(Type_Vehicule As Long, Field1 As String) As String

'Field1 = Nom du field de la case NEUVE

Select Case Field1
Case "Neuve"
    If Type_veh = 0 Then   'type de vehicule
        Field1 = "X"
    Else
        If Type_Vehicule = 1 Then
            Field1 = ""
        Else
            Field1 = ""
        End If
    End If
Case "Usagee"
    If Type_veh = 0 Then   'type de vehicule
        Field1 = ""
    Else
        If Type_Vehicule = 1 Then
            Field1 = "X"
        Else
            Field1 = ""
        End If
    End If
Case "Essai"
    If Type_veh = 0 Then   'type de vehicule
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




Public Function TrouverProduit_montant(ID_vehicule As Long, Piece As Long, Cat_Field As String) As Variant
Select Case Cat_Field
Case "X"
    If TrouverProduit(ID_vehicule, Piece) > -1 Then
        TrouverProduit_montant = "X"
    End If
Case "Montant"
    If TrouverProduit(ID_vehicule, Piece) > -1 Then
        If Montant > 0 Then
            TrouverProduit_montant = Format(TrouverProduit(ID_vehicule, Piece), "#,##0.00")
        End If
    End If
End Select
End Function
Public Function TrouverProduit(ID_vehicule As Long, ID_Produit As Long) As Long
Dim rstTemp1 As ADODB.Recordset
Dim sSql9 As String
Set rstTemp1 = New ADODB.Recordset
    
sSql9 = "SELECT ASINTVEHPRO.IDINTVEH, ASINTVEHPRO.IDPRO, ASINTVEHPRO.MONTANT " & _
        "From ASINTVEHPRO " & _
        "WHERE (((ASINTVEHPRO.IDINTVEH)= " & ID_vehicule & " ) AND ((ASINTVEHPRO.IDPRO)= " & ID_Produit & " ));"

rstTemp1.Open sSql9, oConnection

If Not (rstTemp1.EOF) Then
    TrouverProduit = rstTemp1!Montant
Else
    TrouverProduit = -1
End If

If rstTemp1.State = adStateOpen Then rstTemp1.Close
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
    
    rstTemp1.Open sSql10, oConnection
    
    Do Until rstTemp1.EOF
        strListe = strListe & rstTemp1!DESC0
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
            
    rstTemp2.Open sSql5, oConnection

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

Public Function Garantie(Annee As Long, Km As Long, Type_veh As Long, Lettre_case As String) As Boolean

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

Garantie = flag

End Function
