Private Sub cmdXMLIntoMSXML_Click()
Dim xmlDoc As MSXML.DOMDocument
Dim objNodeList As MSXML.IXMLDOMNodeList
Dim objNode As MSXML.IXMLDOMNode, i As Integer
Dim objAttribute As MSXML.IXMLDOMAttribute
Dim objNodeMap As MSXML.IXMLDOMNamedNodeMap
Dim objNamedItem As MSXML.IXMLDOMNode
Dim sXML As String
Dim sNodeToFind As String
    Set objPub = New Publication.Title
    Set xmlDoc = New MSXML.DOMDocument
    sXML = objPub.RetrieveTitlesSimpleXML()
    sNodeToFind = "Book"
    xmlDoc.async = False
    xmlDoc.loadXML (sXML)
    Set objNodeList = xmlDoc.getElementsByTagName(sNodeToFind)
    For i = 0 To (objNodeList.length - 1)
      Set objNode = objNodeList.nextNode
      
      Set objNodeMap = objNode.Attributes
      Set objNamedItem = objNodeMap.getNamedItem("TitleID")
      txtNotes = txtNotes & objNamedItem.Text & "  -  "
      txtNotes = txtNotes & objNode.Text & vbCrLf
    Next
    Set objPub = Nothing
    Set xmlDoc = Nothing
End Sub
