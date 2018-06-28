Attribute VB_Name = "Module3"
Function RetrieveTitlesSimpleXML() As Variant
Dim vOutPut As Variant, sReturn As String
Dim i As Integer
Dim sTitle As String, sNotes As String
Dim sBeginTag As String
    Set oXML = New Generator
    Set rsTitle = New ADODB.Recordset
    vOutPut = oXML.XMLDeclaration()
    vOutPut = vOutPut & oXML.BeginTag("Books")
    sSQL = "SELECT * FROM titles"
    rsTitle.Open sSQL, GetDSN, adOpenForwardOnly, ­_
    adLockReadOnly , adCmdText
        
    Do While Not rsTitle.EOF
        sTitle = oXML.Format("Title", rsTitle("title"))
        sNotes = oXML.Format("Notes", rsTitle("notes"))
        sBeginTag = oXML.BeginTag("Book", "TitleID", _
                rsTitle("title_id"))
        vOutPut = vOutPut & sBeginTag & sTitle & sNotes _
                & oXML.EndTag("Book")
        rsTitle.MoveNext
    Loop
    vOutPut = vOutPut & oXML.EndTag("Books")
    rsTitle.Close
    Set rsTitle = Nothing
    RetrieveTitlesSimpleXML = vOutPut
End Function
