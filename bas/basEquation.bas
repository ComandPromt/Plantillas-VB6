Attribute VB_Name = "basEquation"
Option Explicit

'Error defines for clsEquation

Public Const EQ_PAREN = 1100     ' Unbalanced parenthesis
Public Const EQ_FUNCTION = 1101  ' Unknown function:
Public Const EQ_VARIABLE = 1102  ' Unknown variable:
Public Const EQ_INVALID = 1103   ' Invalid Equation
Public Const EQ_ARGS = 1104      ' Invalids arguments to function:
Public Const EQ_NAME = 1105      ' Unable to add an unnamed function:
