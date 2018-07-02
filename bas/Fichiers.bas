Attribute VB_Name = "Fichiers"
Option Explicit

Public Function ChercheTitre(sFic As String) As String
Dim i As Integer
Dim st As String * 256
Dim sTitre As String
Dim sf As String * 4
Open sFic For Binary As #1
    Get #1, 1, st
    Get #1, 23, sf
Close
i = InStr(1, st, Chr(0) & Chr(255) & Chr(3))
If i = 0 Then
    i = InStr(1, st, Chr(0) & Chr(255) & Chr(6))
End If
If i Then
    sTitre = Mid(st, i + 4, Asc(Mid(st, i + 3, 1)))
    Debug.Print sTitre
End If
''If InStr(1, st, Chr(0)) > 0 Then ChercheTitre = Left(st, InStr(1, st, Chr(0)) - 1)
'For i = 1 To 4
'    Debug.Print Asc(Mid(sf, i, 1));
'Next
'Debug.Print
'Debug.Print st
'For i = 1 To 100
'    Debug.Print Asc(Mid(st, i, 1)); Mid(st, i, 1);
'Next
ChercheTitre = sTitre
End Function

Public Sub ChercheTitres(s As String, l As ListBox)
Dim i As Integer
Dim sc As String * 1
Dim scc As String * 1
Dim scl As String * 1
Dim sTitre As String
Open s For Binary As #1
    l.Clear
    Do
        Get #1, , sc
        If sc = Chr(0) Then
            Get #1, , sc
            If sc = Chr(255) Then
                Get #1, , scc
                'If scc = Chr(3) Or scc = Chr(6) Then
                    Get #1, , scl
                    sTitre = ""
                    For i = 1 To Asc(scl)
                        Get #1, , sc
                        sTitre = sTitre & sc
                    Next
                    l.AddItem sTitre
                    DoEvents
                'End If
            End If
        End If
    Loop Until EOF(1)
Close
End Sub

Public Function sIsoleRep(sFic As String) As String
    Dim i As Integer
    
    For i = Len(sFic) To 1 Step -1
        If Mid(sFic, i, 1) = "\" Then Exit For
    Next
    If i > 0 Then sIsoleRep = Left(sFic, i - 1)
End Function

'²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²
' Rep : répertoire à explorer
' Fic : tableau description des fichiers
' N : Nb de fichiers explorés (de 0 à N-1)
'²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²²
Public Sub LitFichiers(ByRef N As Integer, Fic() As String, Rep As String)
    Dim s As String
    
    ' Extrait la première entrée.
    s = Dir(Rep & "\*.*", vbDirectory)
    Do While s <> ""  ' Commence la boucle.
        If s <> "." And s <> ".." And GetAttr(s) <> vbDirectory And InStr(1, s, ".mid", vbTextCompare) <> 0 Then ' Ignore les répertoires et les fichiers non MIDI.
            ReDim Preserve Fic(N)
            Fic(N) = s
            N = N + 1
        End If
        s = Dir   ' Extrait l'entrée suivante.
    Loop
End Sub

