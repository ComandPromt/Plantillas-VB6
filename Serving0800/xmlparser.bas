Option Explicit
Const xs = "<"
Const xe = ">"
Const xend = "</"

Public Function XMLDeclaration() As String
  XMLDeclaration = "<?xml version=""1.0""?>"
End Function

Public Function Format(XMLName As String, XMLValue As Variant) As String

    Select Case VarType(XMLValue)
    Case vbByte, vbInteger, vbSingle, vbDouble, vbDecimal, vbBoolean, vbLong, vbCurrency
        Format = xs + XMLName + xe + Trim(Str(XMLValue)) + xend + XMLName + xe
    Case vbString, vbVariant, vbDate
        Format = xs + XMLName + xe + XMLValue + xend + XMLName + xe
    Case Else
        Format = ""
    End Select

End Function

Public Function BeginTag(XMLName As String, _
    Optional AttributeName As String = "", Optional AttributeData As String = "")
Dim sTempTag As String
    sTempTag = xs & XMLName
    
    If AttributeName <> "" Then
        sTempTag = sTempTag & " " & AttributeName & "=" & Chr(34) & AttributeData & """"
    End If
    
    BeginTag = sTempTag & xe
End Function

Public Function EndTag(XMLName As String)
    EndTag = xend & XMLName & xe
End Function
