Attribute VB_Name = "ModuleComplexMath"
'************************************************************************************************
'The Following SubRoutines Are Found In This Module
'------------------------------------------------------------------------------------------------
'ComplexAddition(Real1, Imaginary1, Real2, Imaginary2, RealAnswer, ImaginaryAnswer)
'ComplexSubtraction(Real1, Imaginary1, Real2, Imaginary2, RealAnswer, ImaginaryAnswer)
'ComplexCosine(Real, Imaginary, RealAnswer, ImaginaryAnswer)
'ComplexDivision(Real1, Imaginary1, Real2, Imaginary2, RealAnswer, ImaginaryAnswer)
'ComplexExponentiation(Real, Imaginary, RealAnswer, ImaginaryAnswer)
'ComplexMagnitude(Real, Imaginary, RealAnswer)
'ComplexMultiplication(Real1, Imaginary1, Real2, Imaginary2, RealAnswer, ImaginaryAnswer)
'ComplexSine(Real1, Imaginary1, RealAnswer, ImaginaryAnswer)
'ComplexTangent(Real1, Imaginary1, RealAnswer, ImaginaryAnswer)
'ComplexLogWithSpecialBase(Real, Imaginary, RealBase, ImaginaryBase, RealAnswer, ImaginaryAnswer)
'ComplexLog(Real, Imaginary, RealAnswer, ImaginaryAnswer)
'ComplexPower(Real1, Imaginary1, Real2, Imaginary2, RealAnswer, ImaginaryAnswer)
'------------------------------------------------------------------------------------------------
'Great Module To Throw Into Any Program Requiring Complex Number Manipulation,
'Especially A Fractal Program, Since Almost Every Complex Function Is Here.
'------------------------------------------------------------------------------------------------
'Please Do Not Modify This Modules's Code. It All Works Quite Nicely.
'Visit My HomePage At http://members.aol.com/soze99 For Some GREAT Looking Fractals.
'Also, Please E-Mail Any Improvements To This Program To Soze99@aol.com
'Thanks,
'Soze99 - 1/14/00
'************************************************************************************************
Option Explicit
Dim pi As Double
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim suma As Double
Dim sumb As Double
Dim diffa As Double
Dim diffb As Double
Dim AnswerA As Double
Dim AnswerB As Double
Dim FirstMultA As Double
Dim FirstMultB As Double
Dim FirsteA As Double
Dim FirsteB As Double
Dim SecondMultA As Double
Dim SecondMultB As Double
Dim SecondeA As Double
Dim SecondeB As Double
Dim SumeA As Double
Dim SumeB As Double
Dim DivisionA As Double
Dim DivisionB As Double
Dim multiplicationa As Double
Dim multiplicationb As Double
Dim Divisor As Double
Dim Magnitude As Double
Dim SineA As Double
Dim SineB As Double
Dim CosineA As Double
Dim CosineB As Double
Dim Answer1A As Double
Dim Answer1B As Double
Dim Answer2A As Double
Dim Answer2B As Double
Dim BaseA As Double
Dim BaseB As Double
Dim LogA As Double
Dim LogB As Double
Dim MultA As Double
Dim MultB As Double

Public Sub ComplexAddition(a, b, c, d, suma, sumb)
    suma = a + c
    sumb = b + d
End Sub
Public Sub ComplexSubtraction(a, b, c, d, diffa, diffb)
    diffa = a - c
    diffb = b - d
End Sub
Public Sub ComplexCosine(a, b, AnswerA, AnswerB)
    Call ComplexMultiplication(a, b, 0, 1, FirstMultA, FirstMultB)
    Call ComplexExponentiation(FirstMultA, FirstMultB, FirsteA, FirsteB)
    Call ComplexMultiplication(a, b, 0, -1, SecondMultA, SecondMultB)
    Call ComplexExponentiation(SecondMultA, SecondMultB, SecondeA, SecondeB)
    Call ComplexAddition(FirsteA, FirsteB, SecondeA, SecondeB, SumeA, SumeB)
    Call ComplexDivision(SumeA, SumeB, 2, 0, AnswerA, AnswerB)
End Sub
Public Sub ComplexDivision(a, b, c, d, DivisionA, DivisionB)
    Call ComplexMultiplication(a, b, c, -d, multiplicationa, multiplicationb)
    Divisor = c ^ 2 + d ^ 2
    DivisionA = multiplicationa / Divisor
    DivisionB = multiplicationb / Divisor
End Sub
Public Sub ComplexExponentiation(a, b, AnswerA, AnswerB)
    AnswerA = Exp(a) * Cos(b)
    AnswerB = Exp(a) * Sin(b)
End Sub
Public Sub ComplexMagnitude(a, b, Magnitude)
    Magnitude = Sqr(a ^ 2 + b ^ 2)
End Sub
Public Sub ComplexMultiplication(a, b, c, d, AnswerA, AnswerB)
    AnswerA = a * c - b * d
    AnswerB = a * d + b * c
End Sub
Public Sub ComplexSine(a, b, AnswerA, AnswerB)
    Call ComplexMultiplication(a, b, 0, 1, FirstMultA, FirstMultB)
    Call ComplexExponentiation(FirstMultA, FirstMultB, FirsteA, FirsteB)
    Call ComplexMultiplication(a, b, 0, -1, SecondMultA, SecondMultB)
    Call ComplexExponentiation(SecondMultA, SecondMultB, SecondeA, SecondeB)
    Call ComplexAddition(FirsteA, FirsteB, -SecondeA, -SecondeB, SumeA, SumeB)
    Call ComplexDivision(SumeA, SumeB, 2, 1, AnswerA, AnswerB)
End Sub
Public Sub ComplexTangent(a, b, AnswerA, AnswerB)
    Call ComplexSine(a, b, SineA, SineB)
    Call ComplexCosine(a, b, CosineA, CosineB)
    Call ComplexDivision(SineA, SineB, CosineA, CosineB, AnswerA, AnswerB)
End Sub
Public Sub ComplexLogWithSpecialBase(a, b, BaseA, BaseB, AnswerA, AnswerB)
    Call ComplexLog(a, b, Answer1A, Answer1B)
    Call ComplexLog(BaseA, BaseB, Answer2A, Answer2B)
    Call ComplexDivision(Answer1A, Answer1B, Answer2A, Answer2B, AnswerA, AnswerB)
End Sub
Public Sub ComplexLog(a, b, AnswerA, AnswerB)
    pi = 4 * Atn(1)
    Call ComplexMagnitude(a, b, Magnitude)
    If Magnitude = 0 Then
        AnswerA = 0
    Else
        AnswerA = Log(Magnitude)
    End If
    If (a >= 0 And b = 0) Or (a = 0 And b = 0) Then
        AnswerB = 0
    ElseIf a = 0 And b >= 0 Then
        AnswerB = pi / 2
    ElseIf a <= 0 And b = 0 Then
        AnswerB = pi
    ElseIf a = 0 And b <= 0 Then
        AnswerB = 3 * pi / 2
    Else
        AnswerB = pi / 2 * (-a / Abs(a) + 1) + pi * a / Abs(a) * (-b / Abs(b) + 1) + a * b / Abs(a * b) * Atn(Abs(b / a))
    End If
End Sub
Public Sub ComplexPower(a, b, c, d, AnswerA, AnswerB)
    Call ComplexLog(a, b, LogA, LogB)
    Call ComplexMultiplication(c, d, LogA, LogB, MultA, MultB)
    Call ComplexExponentiation(MultA, MultB, AnswerA, AnswerB)
End Sub
