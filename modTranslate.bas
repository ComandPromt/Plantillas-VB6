Attribute VB_Name = "modTranslate"
'------------------------------------------
'
' Walther Musch
' 30041999
'
' http://www.kather.net/VisualBasicSource
'
'------------------------------------------
 
Option Explicit
'
Private colTranslation As Collection
Private vWords()

Public Sub SetDefault()
    With frmTranslate
        .Caption = "Simple Translation"
        .lblText.Caption = "Type some words in english and click on Translate." + vbCrLf + _
            "You will get a translation of known words in dutch."
        .txtInput.Text = "Well, how are you today?"
    End With
    Call InitCollection
End Sub

Public Sub InitCollection()
    Set colTranslation = New Collection
    '
    'the format of the collection is
    '[value],[key]
    'for this purpose the [value] is the translated word
    'while [key] is the word to be translated
    '
    Call colTranslation.Add("hallo", "well")
    Call colTranslation.Add("hoe", "how")
    Call colTranslation.Add("jij", "you")
    Call colTranslation.Add("vandaag", "today")
    '
    'you can fill the collection also from a file .....
End Sub

Public Sub Translating()
    Dim intX As Integer
    ''
    'first split the string into words
    Call SplitStringintoWords(frmTranslate.txtInput.Text)
    'second translate one by one and show result
    frmTranslate.txtInput.Text = frmTranslate.txtInput.Text + vbCrLf
    For intX = LBound(vWords) To UBound(vWords)
        'MsgBox ExistInCollectionString(colTranslation, CStr(vWords(intX)))
        frmTranslate.txtInput.Text = frmTranslate.txtInput.Text + ExistInCollectionString(colTranslation, CStr(vWords(intX))) + " "
    Next intX
End Sub

Private Sub SplitStringintoWords(ByRef strSource As String)
    'splitting the inputstrin into words
    'and save then in a array
    '
    Dim intX As Integer
    Dim intY As Integer
    Dim intZ As Integer
    Dim TempstrSource As String
    Dim tmp As String
    '
    intZ = 0
    TempstrSource = strSource
    For intX = 1 To Len(strSource)
        intY = InStr(TempstrSource, Chr(32))
        If intY <> 0 Then
            ReDim Preserve vWords(intZ)
            tmp = Left$(TempstrSource, intY - 1)
            vWords(intZ) = StripString(tmp)
            TempstrSource = Right$(TempstrSource, Len(TempstrSource) - intY)
            intZ = intZ + 1
            intX = intX + intY
        End If
    Next intX
    ReDim Preserve vWords(intZ)
    vWords(intZ) = StripString(TempstrSource)

End Sub

Private Function StripString(ByRef strSource As String) As String
    'remove all other characters but letters
    '
    Const Letters As String = "abcdefghijklmnopqrstuvwxyz"
    Dim intX As Integer
    Dim tmp As String
    '
    tmp = strSource
    For intX = 1 To Len(strSource)
        If InStr(Letters, LCase(Mid$(strSource, intX, 1))) = 0 Then
            Select Case intX
            Case 1
                tmp = Right$(strSource, Len(strSource) - intX)
            Case Len(strSource)
                tmp = Left$(strSource, Len(strSource) - 1)
            Case Else
                tmp = Left$(strSource, intX) & Right$(strSource, Len(strSource) - intX)
            End Select
        End If
    Next intX
    StripString = tmp
            
End Function

Private Function ExistInCollectionString(colItems As Collection, strItem As String) As String
    'search for the [key] (the word to be translated) in the collection
    'if found return the [value]
    'if not the code resume at the label 'ExistInCollectionString_Fail'
    '
    On Error GoTo ExistInCollectionString_Fail
    ExistInCollectionString = colItems.Item(strItem)
    Exit Function
    '
ExistInCollectionString_Fail:
    ExistInCollectionString = "?[notfound]?"
End Function

