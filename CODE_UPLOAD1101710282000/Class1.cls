VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExpressionParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const PI = 3.14159265358979
Private Const GENERIC_SYNTAX_ERR_MSG = "Syntax Error"
Public Enum ParserErrors
    PERR_FIRST = vbObjectError + 513
    PERR_SYNTAX_ERROR = PERR_FIRST
    PERR_DIVISION_BY_ZERO
    PERR_CLOSING_PARENTHESES_EXPECTED
    PERR_INVALID_CONST_NAME
    PERR_FUNCTION_DOES_NOT_EXIST
    PERR_RESERVED_WORD
    PERR_CONST_ALREADY_EXISTS
    PERR_CONST_DOES_NOT_EXIST
    PERR_LAST = PERR_CONST_DOES_NOT_EXIST
End Enum

Private Enum ParserTokens
    TOK_UNKNOWN
    TOK_FIRST
    TOK_ADD = TOK_FIRST
    TOK_SUBTRACT
    TOK_MULTIPLY
    TOK_DIVIDE
    TOK_OPEN_PARENTHESES
    TOK_CLOSE_PARENTHESES
    TOK_LAST = TOK_CLOSE_PARENTHESES
End Enum
Private mTokenSymbols() As String

Private mExpression As String
Private mPosition As Long
Private mLastTokenLength As Long
Private mConstants As Collection
Private mProjectName As String
Public Function ParseExpression(Expression As String) As Double

On Error GoTo ParseExpression_ErrHandler

Dim Value As Double

    mExpression = Expression
    mPosition = 1
    
    SkipSpaces
    Value = ParseNumExp
    SkipSpaces
    
    If mPosition <= Len(mExpression) Then
        Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
    End If

    ParseExpression = Value
    Exit Function
    
ParseExpression_ErrHandler:
    SetErrSource "ParseExpression"
    Err.Raise Err.Number
End Function
Private Function ParseNumExp() As Double
On Error GoTo ParseNumExp_ErrHandler

Dim Value As Double
Dim OtherValue As Double
Dim CurrToken As ParserTokens
    
    Value = ParseTerm
        
    Do While mPosition <= Len(mExpression)
       
        CurrToken = GetToken

        If CurrToken = TOK_ADD Then

            SkipLastToken
            OtherValue = ParseTerm
            Value = Value + OtherValue
        ElseIf CurrToken = TOK_SUBTRACT Then
            SkipLastToken
            OtherValue = ParseTerm
            Value = Value - OtherValue
        ElseIf CurrToken = TOK_UNKNOWN Then
            Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
        Else

            ParseNumExp = Value
            Exit Function
        End If
    Loop

    If Not IsNull(Value) Then
        ParseNumExp = Value
    Else
        ParseNumExp = 0
    End If
    Exit Function
    
ParseNumExp_ErrHandler:
    SetErrSource "ParseNumExp"
    Err.Raise Err.Number
End Function
Private Function ParseTerm() As Double
On Error GoTo ParseTerm_ErrHandler

Dim Value As Double
Dim OtherValue As Double
Dim CurrToken As ParserTokens

    Value = ParseValue
    
    Do While mPosition <= Len(mExpression)
        
        CurrToken = GetToken

        If CurrToken = TOK_MULTIPLY Then
            SkipLastToken
            
            OtherValue = ParseValue
            Value = Value * OtherValue
        ElseIf CurrToken = TOK_DIVIDE Then
            SkipLastToken
            
            OtherValue = ParseValue
            If OtherValue = 0 Then
                Err.Raise PERR_DIVISION_BY_ZERO, , _
                    "Division by Zero!"
            End If
            
            Value = Value / OtherValue
        ElseIf CurrToken = TOK_UNKNOWN Then
            Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
        Else
            ParseTerm = Value
            Exit Function
        End If
    Loop

    ParseTerm = Value

    Exit Function
    
ParseTerm_ErrHandler:
    SetErrSource "ParseTerm"
    Err.Raise Err.Number
End Function
Private Function ParseValue() As Double
On Error GoTo ParseValue_ErrHandler

Dim Sign As Double
Dim CurrToken As ParserTokens
Dim Value As Double
Dim IsValue As Boolean

    Sign = 1

    CurrToken = GetToken
    If CurrToken = TOK_SUBTRACT Then
        Sign = -1
        SkipLastToken
    ElseIf CurrToken = TOK_ADD Then

        SkipLastToken
    End If

    CurrToken = GetToken
    If CurrToken = TOK_OPEN_PARENTHESES Then
        SkipLastToken
        Value = ParseNumExp
        CurrToken = GetToken
        If CurrToken = TOK_CLOSE_PARENTHESES Then
            SkipLastToken
        Else
            Err.Raise PERR_CLOSING_PARENTHESES_EXPECTED, , "')' Expected"
        End If
    Else
        Value = ParseAtom
    End If

    ParseValue = Value * Sign
    Exit Function
    
ParseValue_ErrHandler:
    SetErrSource "ParseValue"
    Err.Raise Err.Number
End Function
Private Function ParseAtom() As Double
On Error GoTo ParseAtom_ErrHandler

Dim CurrPosition As Long
Dim CurrToken As ParserTokens
Dim SymbolName As String
Dim ArgumentValue As Double
Dim DecimalPointFound As Boolean
Dim Value As Double
Dim IsValue As Boolean

    If mPosition > Len(mExpression) Then
        Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
    End If

    CurrPosition = mPosition
    
    IsValue = False
    
    If IsNumeric(Mid(mExpression, CurrPosition, 1)) Then
        IsValue = True
        
        CurrPosition = CurrPosition + 1
        DecimalPointFound = False
        
        ' Read the rest of the number
        Do While IsNumeric(Mid(mExpression, CurrPosition, 1)) Or _
            Mid(mExpression, CurrPosition, 1) = "."
            
            If Mid(mExpression, CurrPosition, 1) = "." Then
                If Not DecimalPointFound Then
                    DecimalPointFound = True
                Else
                    Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
                End If
            End If
            
            CurrPosition = CurrPosition + 1
        Loop
        
        Value = CDbl(Mid(mExpression, mPosition, CurrPosition - mPosition))
        
        mPosition = CurrPosition
        SkipSpaces
    End If

    If Not IsValue Then
        If IsLetter(Mid(mExpression, CurrPosition, 1)) Then
            CurrPosition = CurrPosition + 1
            
            Do While CurrPosition <= Len(mExpression)
                If IsValidSymbolCharacter(Mid(mExpression, CurrPosition, 1)) Then
                    CurrPosition = CurrPosition + 1
                Else
                    Exit Do
                End If
            Loop
                    
            SymbolName = Mid(mExpression, mPosition, CurrPosition - mPosition)
            mPosition = CurrPosition
            SkipSpaces
        
            CurrToken = GetToken
            If CurrToken = TOK_OPEN_PARENTHESES Then
                SkipLastToken
                ArgumentValue = ParseNumExp
                
                CurrToken = GetToken
                If CurrToken = TOK_CLOSE_PARENTHESES Then
                    SkipLastToken
                Else
                    Err.Raise PERR_CLOSING_PARENTHESES_EXPECTED, , "')' Expected"
                End If
            
                Value = CallBuiltinFunction(SymbolName, _
                    ArgumentValue)
                IsValue = True
            Else
                If ConstExists(SymbolName) Then
                    Value = mConstants(SymbolName)
                    IsValue = True
                Else
                    Err.Raise PERR_CONST_DOES_NOT_EXIST, , _
                        "Constant name " & SymbolName & " does not exist"
                End If
            End If
        End If
    End If

    If Not IsValue Then
        Err.Raise PERR_SYNTAX_ERROR, , GENERIC_SYNTAX_ERR_MSG
    End If

    ParseAtom = Value
    Exit Function
    
ParseAtom_ErrHandler:
    SetErrSource "ParseAtom"
    Err.Raise Err.Number
End Function

Private Function GetToken() As ParserTokens
Dim CurrToken As ParserTokens
Dim i As ParserTokens
    
    If mPosition > Len(mExpression) Then
        GetToken = TOK_UNKNOWN
        Exit Function
    End If
    
    CurrToken = TOK_UNKNOWN
    mLastTokenLength = 0
    
    For i = TOK_FIRST To TOK_LAST
        If Mid(mExpression, mPosition, Len(mTokenSymbols(i))) = mTokenSymbols(i) Then
            CurrToken = i
            mLastTokenLength = Len(mTokenSymbols(i))
            Exit For
        End If
    Next
    
    GetToken = CurrToken
End Function

Private Sub SkipLastToken()
    
    mPosition = mPosition + mLastTokenLength
    SkipSpaces

End Sub
Private Function ConstExists(Name As String) As Boolean
Const ERR_KEY_NOT_FOUND = 5
Dim DummyValue As Double

    On Error Resume Next
    DummyValue = mConstants(Name)
    
    If Err.Number = ERR_KEY_NOT_FOUND Then
        ConstExists = False
    Else
        ConstExists = True
    End If
    
End Function

Public Sub AddConstant(Name As String, Value As Double)
Dim i As ParserTokens
Dim TempName As String

    TempName = UCase(Trim(Name))
   
    If Len(TempName) = 0 Then
        Err.Raise PERR_INVALID_CONST_NAME, , "Constant name cannot be null"
    End If

    If Not IsLetter(Left(TempName, 1)) Then
        Err.Raise PERR_INVALID_CONST_NAME, , "Constant name must begin with a letter"
    End If
            
    For i = 2 To Len(TempName)
        If Not IsValidSymbolCharacter(Mid(TempName, i, 1)) Then
            Err.Raise PERR_INVALID_CONST_NAME, , "Invalid constant name"
        End If
    Next
    
    If ConstExists(TempName) Then
        Err.Raise PERR_CONST_ALREADY_EXISTS, , "The constant already exists"
    End If
    
    If IsBuiltInFunction(TempName) Then
        Err.Raise PERR_RESERVED_WORD, , "The name is a reserved word"
    End If
    
    mConstants.Add Value, TempName

End Sub

Public Sub RemoveConstant(Name As String)
Dim TempName As String

    TempName = UCase(Trim(Name))
    
    If ConstExists(TempName) Then
        mConstants.Remove TempName
    Else
        Err.Raise PERR_CONST_DOES_NOT_EXIST, , _
            "Constant name " & TempName & " does not exist"
    End If
    
End Sub
Private Function IsBuiltInFunction(Name As String) As Boolean
Dim TempName As String

    TempName = UCase(Trim(Name))
    If TempName = "SIN" Or _
       TempName = "COS" Or _
       TempName = "ABS" Then

        IsBuiltInFunction = True
    Else
        IsBuiltInFunction = False
    End If

End Function
Public Function IKiif(Champ As Boolean, Si_vrai As Variant, Sinon As Variant) As Variant
IKiif = IIf((Champ = True), Si_vrai, Sinon)

End Function
Private Function CallBuiltinFunction(Name As String, Argument As Double) As Double
On Error GoTo CallBuiltinFunction_ErrHandler

Const DEGREES_TO_RADIANS = PI / 180
Dim TempName As String

    If Not IsBuiltInFunction(Name) Then
        Err.Raise PERR_FUNCTION_DOES_NOT_EXIST, , _
            "Function " & Name & " Does not exist"
    End If
    
    TempName = UCase(Trim(Name))

    Select Case TempName
    Case "SIN"
         CallBuiltinFunction = Sin(Argument * DEGREES_TO_RADIANS)
    Case "COS"
         CallBuiltinFunction = Cos(Argument * DEGREES_TO_RADIANS)
    Case "ABS"
         CallBuiltinFunction = Abs(Argument)
    End Select

    Exit Function
    
CallBuiltinFunction_ErrHandler:
    SetErrSource "CallBuiltinFunction"
    Err.Raise Err.Number
End Function

Private Sub SkipSpaces()
    
    Do While mPosition <= Len(mExpression) And _
             (Mid(mExpression, mPosition, 1) = " " Or _
              Mid(mExpression, mPosition, 1) = vbTab)
        mPosition = mPosition + 1
    Loop

End Sub

Private Function IsLetter(Character As String) As Boolean
Dim CharAsciiCode As Long

    CharAsciiCode = Asc(UCase(Character))
    If (CharAsciiCode >= Asc("A") And _
        CharAsciiCode <= Asc("Z")) Then
        IsLetter = True
    Else
        IsLetter = False
    End If

End Function
Private Function IsValidSymbolCharacter(Character As String) As Boolean

    If IsLetter(Character) Or _
       IsNumeric(Character) Or _
       Character = "_" Then
            
        IsValidSymbolCharacter = True
    Else
        IsValidSymbolCharacter = False
    End If

End Function
Private Function GetProjectName() As String
On Error Resume Next

    ' Quite a way to get the project name...
    Err.Raise 1, , " "
    GetProjectName = Err.Source
    Err.Clear

End Function

Private Sub SetErrSource(Name As String)
    
    If Err.Source = mProjectName Then
        Err.Source = Name
    Else
        Err.Source = Name & "->" & Err.Source
    End If

End Sub
Public Property Get LastErrorPosition() As Long
    LastErrorPosition = mPosition
End Property

Private Sub Class_Initialize()
    
    ReDim mTokenSymbols(TOK_FIRST To TOK_LAST)

    mTokenSymbols(TOK_ADD) = "+"
    mTokenSymbols(TOK_SUBTRACT) = "-"
    mTokenSymbols(TOK_MULTIPLY) = "*"
    mTokenSymbols(TOK_DIVIDE) = "/"
    mTokenSymbols(TOK_OPEN_PARENTHESES) = "("
    mTokenSymbols(TOK_CLOSE_PARENTHESES) = ")"
    Set mConstants = New Collection
    mConstants.Add PI, "PI"
    
    mProjectName = GetProjectName

End Sub
Public Function exec_fonction_matrice(Object_Name As Object, Fonction_Name As String, matrice() As Variant) As Variant
Dim i As Integer
Dim temp As Variant

'If Len(matrice()) Then

Select Case UBound(matrice())
    Case 0
        exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, matrice(0))
    Case 1
        exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, CStr(matrice(0)), CStr(matrice(1)))
    Case 2
        exec_fonction_matrice = CallByName(Object_Name, Fonction_Name, VbMethod, Me.ParseExpression(CStr(matrice(0))), matrice(1), matrice(2))
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
MsgBox exec_fonction_matrice
End Function

Public Function exec_fonction(Object_Name As Project1.clsExpressionParser, Fonction_Name As String, texte As String) As Variant
Dim i As Integer
Dim matrice() As String
matrice() = Split(texte, ",")

'If Len(matrice()) Then

Select Case UBound(matrice())
    Case 0
        exec_fonction = CallByName(Object_Name, Fonction_Name, VbMethod, (0))
    Case 1
        exec_fonction = CallByName(Object_Name, Fonction_Name, VbMethod, , CStr(matrice(1)), "")
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
MsgBox exec_fonction
End Function
