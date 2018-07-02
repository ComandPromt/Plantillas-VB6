Attribute VB_Name = "Log"
Public num_fichier_log As Integer

Public Sub ouvrir_log()
    ' ouverture du fichier log
    num_fichier_log = FreeFile
    Open ".\Crack80.log" For Append As #num_fichier_log
End Sub

Public Sub fermer_log()
    Close #num_fichier_log
End Sub

Public Sub ecrire_log(message As String)
    Print #num_fichier_log, message
End Sub
