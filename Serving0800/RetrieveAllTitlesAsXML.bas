Attribute VB_Name = "Module2"
Function RetrieveAllTitlesAsXML() As Variant
Dim stXMLStream As Stream 'dimensions the stXMLStream variable as a Stream object
    Set rsAllTitles = New ADODB.Recordset 'creates a new instance of the ADO Recordset object
    Set stXMLStream = New Stream 'creates a new instance of the Stream object
    sSQL = "select * from titles" 'sets the SQL for the Open method
    rsAllTitles.Open sSQL, GetDSN, adOpenForwardOnly, adLockReadOnly, adCmdText 'executes Open to retrieve the records
    rsAllTitles.save stXMLStream, adPersistXML 'saves the recordset to the Stream object in XML format by specifying the stXMLStream variable as the first argument to the Save method then specifying the adPersistXML constant as the last argument
    rsAllTitles.Close 'close the recordset
    Set rsAllTitles = Nothing 'set the recordset variable to "Nothing"
    Set RetrieveAllTitlesAsXML = stXMLStream 'returns the Stream object as the return variable from the function
'At this point we have access to the Stream object
End Function

Function stXMLStream()
Dim stXMLStream As Stream 'creates a variable and defines it as type Stream
  Set objPub = New Publication.Title ' instantiates the Title object
  Set stXMLStream = objPub.RetrieveAllTitlesAsXML() 'executes RetrieveAllTitlesAsXML and sets the return value to the stXMLStream variable
  txtNotes.Text = stXMLStream.ReadText() 'extracts the XML from the Stream object by executing the ReadText method of the Stream object
End Function

Function use()
sXML = stXMLStream.ReadText()
xmlDoc.loadXML (sXML)
Response.Write "<?xml version='1.0' encoding='ISO-8859-1'?>" & vbCrLf
rstCustomers.save Response, adPersistXML
rstCustomers.Close
End Function
