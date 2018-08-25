Attribute VB_Name = "Vowels"
Public VowelCount As Integer
Public Filenumber As Integer
Public Function CountVowels(Vowel As String)
Dim VowelLocation As String
Dim length As Integer
lenght = Len(Vowel)
VowelsInput = UCase(Vowel)
VowelCount = 0
For I = 1 To lenght
VowelLocation = Mid(Vowel, I, 1)
Select Case VowelLocation
Case "a"
VowelCount = VowelCount + 1
Case "A"
VowelCount = VowelCount + 1
Case "e"
VowelCount = VowelCount + 1
Case "E"
VowelCount = VowelCount + 1
Case "i"
VowelCount = VowelCount + 1
Case "I"
VowelCount = VowelCount + 1
Case "o"
VowelCount = VowelCount + 1
Case "O"
VowelCount = VowelCount + 1
Case "u"
VowelCount = VowelCount + 1
Case "U"
VowelCount = VowelCount + 1
End Select
Next I
End Function


