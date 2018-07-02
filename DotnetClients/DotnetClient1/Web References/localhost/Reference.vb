﻿'------------------------------------------------------------------------------
' <autogenerated>
'     This code was generated by a tool.
'     Runtime Version: 1.0.3705.209
'
'     Changes to this file may cause incorrect behavior and will be lost if 
'     the code is regenerated.
' </autogenerated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Diagnostics
Imports System.Runtime.Serialization
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml
Imports System.Xml.Serialization

'
'This source code was auto-generated by Microsoft.VSDesigner, Version 1.0.3705.209.
'
Namespace localhost
    
    '<remarks/>
    <System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Web.Services.WebServiceBindingAttribute(Name:="DataSetServiceSoap", [Namespace]:="http://example.org/dataset-service")>  _
    Public Class DataSetService
        Inherits System.Web.Services.Protocols.SoapHttpClientProtocol
        
        '<remarks/>
        Public Sub New()
            MyBase.New
            Me.Url = "http://localhost/datasetservice/datasetservice.asmx"
        End Sub
        
        '<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://example.org/dataset-service/GetAuthorsAsTypedDataSet", RequestNamespace:="http://example.org/dataset-service", ResponseNamespace:="http://example.org/dataset-service", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function GetAuthorsAsTypedDataSet() As AuthorSet
            Dim results() As Object = Me.Invoke("GetAuthorsAsTypedDataSet", New Object(-1) {})
            Return CType(results(0),AuthorSet)
        End Function
        
        '<remarks/>
        Public Function BeginGetAuthorsAsTypedDataSet(ByVal callback As System.AsyncCallback, ByVal asyncState As Object) As System.IAsyncResult
            Return Me.BeginInvoke("GetAuthorsAsTypedDataSet", New Object(-1) {}, callback, asyncState)
        End Function
        
        '<remarks/>
        Public Function EndGetAuthorsAsTypedDataSet(ByVal asyncResult As System.IAsyncResult) As AuthorSet
            Dim results() As Object = Me.EndInvoke(asyncResult)
            Return CType(results(0),AuthorSet)
        End Function
        
        '<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://example.org/dataset-service/GetAuthorsAsXml", RequestNamespace:="http://example.org/dataset-service", ResponseNamespace:="http://example.org/dataset-service", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function GetAuthorsAsXml() As System.Xml.XmlNode
            Dim results() As Object = Me.Invoke("GetAuthorsAsXml", New Object(-1) {})
            Return CType(results(0),System.Xml.XmlNode)
        End Function
        
        '<remarks/>
        Public Function BeginGetAuthorsAsXml(ByVal callback As System.AsyncCallback, ByVal asyncState As Object) As System.IAsyncResult
            Return Me.BeginInvoke("GetAuthorsAsXml", New Object(-1) {}, callback, asyncState)
        End Function
        
        '<remarks/>
        Public Function EndGetAuthorsAsXml(ByVal asyncResult As System.IAsyncResult) As System.Xml.XmlNode
            Dim results() As Object = Me.EndInvoke(asyncResult)
            Return CType(results(0),System.Xml.XmlNode)
        End Function
        
        '<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://example.org/dataset-service/GetAuthors", RequestNamespace:="http://example.org/dataset-service", ResponseNamespace:="http://example.org/dataset-service", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function GetAuthors() As System.Data.DataSet
            Dim results() As Object = Me.Invoke("GetAuthors", New Object(-1) {})
            Return CType(results(0),System.Data.DataSet)
        End Function
        
        '<remarks/>
        Public Function BeginGetAuthors(ByVal callback As System.AsyncCallback, ByVal asyncState As Object) As System.IAsyncResult
            Return Me.BeginInvoke("GetAuthors", New Object(-1) {}, callback, asyncState)
        End Function
        
        '<remarks/>
        Public Function EndGetAuthors(ByVal asyncResult As System.IAsyncResult) As System.Data.DataSet
            Dim results() As Object = Me.EndInvoke(asyncResult)
            Return CType(results(0),System.Data.DataSet)
        End Function
    End Class
    
    <Serializable(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Diagnostics.DebuggerStepThrough(),  _
     System.ComponentModel.ToolboxItem(true)>  _
    Public Class AuthorSet
        Inherits DataSet
        
        Private tableauthors As authorsDataTable
        
        Public Sub New()
            MyBase.New
            Me.InitClass
            Dim schemaChangedHandler As System.ComponentModel.CollectionChangeEventHandler = AddressOf Me.SchemaChanged
            AddHandler Me.Tables.CollectionChanged, schemaChangedHandler
            AddHandler Me.Relations.CollectionChanged, schemaChangedHandler
        End Sub
        
        Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
            MyBase.New
            Dim strSchema As String = CType(info.GetValue("XmlSchema", GetType(System.String)),String)
            If (Not (strSchema) Is Nothing) Then
                Dim ds As DataSet = New DataSet
                ds.ReadXmlSchema(New XmlTextReader(New System.IO.StringReader(strSchema)))
                If (Not (ds.Tables("authors")) Is Nothing) Then
                    Me.Tables.Add(New authorsDataTable(ds.Tables("authors")))
                End If
                Me.DataSetName = ds.DataSetName
                Me.Prefix = ds.Prefix
                Me.Namespace = ds.Namespace
                Me.Locale = ds.Locale
                Me.CaseSensitive = ds.CaseSensitive
                Me.EnforceConstraints = ds.EnforceConstraints
                Me.Merge(ds, false, System.Data.MissingSchemaAction.Add)
                Me.InitVars
            Else
                Me.InitClass
            End If
            Me.GetSerializationData(info, context)
            Dim schemaChangedHandler As System.ComponentModel.CollectionChangeEventHandler = AddressOf Me.SchemaChanged
            AddHandler Me.Tables.CollectionChanged, schemaChangedHandler
            AddHandler Me.Relations.CollectionChanged, schemaChangedHandler
        End Sub
        
        <System.ComponentModel.Browsable(false),  _
         System.ComponentModel.DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Content)>  _
        Public ReadOnly Property authors As authorsDataTable
            Get
                Return Me.tableauthors
            End Get
        End Property
        
        Public Overrides Function Clone() As DataSet
            Dim cln As AuthorSet = CType(MyBase.Clone,AuthorSet)
            cln.InitVars
            Return cln
        End Function
        
        Protected Overrides Function ShouldSerializeTables() As Boolean
            Return false
        End Function
        
        Protected Overrides Function ShouldSerializeRelations() As Boolean
            Return false
        End Function
        
        Protected Overrides Sub ReadXmlSerializable(ByVal reader As XmlReader)
            Me.Reset
            Dim ds As DataSet = New DataSet
            ds.ReadXml(reader)
            If (Not (ds.Tables("authors")) Is Nothing) Then
                Me.Tables.Add(New authorsDataTable(ds.Tables("authors")))
            End If
            Me.DataSetName = ds.DataSetName
            Me.Prefix = ds.Prefix
            Me.Namespace = ds.Namespace
            Me.Locale = ds.Locale
            Me.CaseSensitive = ds.CaseSensitive
            Me.EnforceConstraints = ds.EnforceConstraints
            Me.Merge(ds, false, System.Data.MissingSchemaAction.Add)
            Me.InitVars
        End Sub
        
        Protected Overrides Function GetSchemaSerializable() As System.Xml.Schema.XmlSchema
            Dim stream As System.IO.MemoryStream = New System.IO.MemoryStream
            Me.WriteXmlSchema(New XmlTextWriter(stream, Nothing))
            stream.Position = 0
            Return System.Xml.Schema.XmlSchema.Read(New XmlTextReader(stream), Nothing)
        End Function
        
        Friend Sub InitVars()
            Me.tableauthors = CType(Me.Tables("authors"),authorsDataTable)
            If (Not (Me.tableauthors) Is Nothing) Then
                Me.tableauthors.InitVars
            End If
        End Sub
        
        Private Sub InitClass()
            Me.DataSetName = "AuthorSet"
            Me.Prefix = ""
            Me.Namespace = "http://example.org/dataset"
            Me.Locale = New System.Globalization.CultureInfo("en-US")
            Me.CaseSensitive = false
            Me.EnforceConstraints = true
            Me.tableauthors = New authorsDataTable
            Me.Tables.Add(Me.tableauthors)
        End Sub
        
        Private Function ShouldSerializeauthors() As Boolean
            Return false
        End Function
        
        Private Sub SchemaChanged(ByVal sender As Object, ByVal e As System.ComponentModel.CollectionChangeEventArgs)
            If (e.Action = System.ComponentModel.CollectionChangeAction.Remove) Then
                Me.InitVars
            End If
        End Sub
        
        Public Delegate Sub authorsRowChangeEventHandler(ByVal sender As Object, ByVal e As authorsRowChangeEvent)
        
        <System.Diagnostics.DebuggerStepThrough()>  _
        Public Class authorsDataTable
            Inherits DataTable
            Implements System.Collections.IEnumerable
            
            Private columnau_id As DataColumn
            
            Private columnau_lname As DataColumn
            
            Private columnau_fname As DataColumn
            
            Private columnphone As DataColumn
            
            Private columnaddress As DataColumn
            
            Private columncity As DataColumn
            
            Private columnstate As DataColumn
            
            Private columnzip As DataColumn
            
            Private columncontract As DataColumn
            
            Friend Sub New()
                MyBase.New("authors")
                Me.InitClass
            End Sub
            
            Friend Sub New(ByVal table As DataTable)
                MyBase.New(table.TableName)
                If (table.CaseSensitive <> table.DataSet.CaseSensitive) Then
                    Me.CaseSensitive = table.CaseSensitive
                End If
                If (table.Locale.ToString <> table.DataSet.Locale.ToString) Then
                    Me.Locale = table.Locale
                End If
                If (table.Namespace <> table.DataSet.Namespace) Then
                    Me.Namespace = table.Namespace
                End If
                Me.Prefix = table.Prefix
                Me.MinimumCapacity = table.MinimumCapacity
                Me.DisplayExpression = table.DisplayExpression
            End Sub
            
            <System.ComponentModel.Browsable(false)>  _
            Public ReadOnly Property Count As Integer
                Get
                    Return Me.Rows.Count
                End Get
            End Property
            
            Friend ReadOnly Property au_idColumn As DataColumn
                Get
                    Return Me.columnau_id
                End Get
            End Property
            
            Friend ReadOnly Property au_lnameColumn As DataColumn
                Get
                    Return Me.columnau_lname
                End Get
            End Property
            
            Friend ReadOnly Property au_fnameColumn As DataColumn
                Get
                    Return Me.columnau_fname
                End Get
            End Property
            
            Friend ReadOnly Property phoneColumn As DataColumn
                Get
                    Return Me.columnphone
                End Get
            End Property
            
            Friend ReadOnly Property addressColumn As DataColumn
                Get
                    Return Me.columnaddress
                End Get
            End Property
            
            Friend ReadOnly Property cityColumn As DataColumn
                Get
                    Return Me.columncity
                End Get
            End Property
            
            Friend ReadOnly Property stateColumn As DataColumn
                Get
                    Return Me.columnstate
                End Get
            End Property
            
            Friend ReadOnly Property zipColumn As DataColumn
                Get
                    Return Me.columnzip
                End Get
            End Property
            
            Friend ReadOnly Property contractColumn As DataColumn
                Get
                    Return Me.columncontract
                End Get
            End Property
            
            Public Default ReadOnly Property Item(ByVal index As Integer) As authorsRow
                Get
                    Return CType(Me.Rows(index),authorsRow)
                End Get
            End Property
            
            Public Event authorsRowChanged As authorsRowChangeEventHandler
            
            Public Event authorsRowChanging As authorsRowChangeEventHandler
            
            Public Event authorsRowDeleted As authorsRowChangeEventHandler
            
            Public Event authorsRowDeleting As authorsRowChangeEventHandler
            
            Public Overloads Sub AddauthorsRow(ByVal row As authorsRow)
                Me.Rows.Add(row)
            End Sub
            
            Public Overloads Function AddauthorsRow(ByVal au_id As String, ByVal au_lname As String, ByVal au_fname As String, ByVal phone As String, ByVal address As String, ByVal city As String, ByVal state As String, ByVal zip As String, ByVal contract As Boolean) As authorsRow
                Dim rowauthorsRow As authorsRow = CType(Me.NewRow,authorsRow)
                rowauthorsRow.ItemArray = New Object() {au_id, au_lname, au_fname, phone, address, city, state, zip, contract}
                Me.Rows.Add(rowauthorsRow)
                Return rowauthorsRow
            End Function
            
            Public Function FindByau_id(ByVal au_id As String) As authorsRow
                Return CType(Me.Rows.Find(New Object() {au_id}),authorsRow)
            End Function
            
            Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
                Return Me.Rows.GetEnumerator
            End Function
            
            Public Overrides Function Clone() As DataTable
                Dim cln As authorsDataTable = CType(MyBase.Clone,authorsDataTable)
                cln.InitVars
                Return cln
            End Function
            
            Protected Overrides Function CreateInstance() As DataTable
                Return New authorsDataTable
            End Function
            
            Friend Sub InitVars()
                Me.columnau_id = Me.Columns("au_id")
                Me.columnau_lname = Me.Columns("au_lname")
                Me.columnau_fname = Me.Columns("au_fname")
                Me.columnphone = Me.Columns("phone")
                Me.columnaddress = Me.Columns("address")
                Me.columncity = Me.Columns("city")
                Me.columnstate = Me.Columns("state")
                Me.columnzip = Me.Columns("zip")
                Me.columncontract = Me.Columns("contract")
            End Sub
            
            Private Sub InitClass()
                Me.columnau_id = New DataColumn("au_id", GetType(System.String), Nothing, System.Data.MappingType.Element)
                Me.Columns.Add(Me.columnau_id)
                Me.columnau_lname = New DataColumn("au_lname", GetType(System.String), Nothing, System.Data.MappingType.Element)
                Me.Columns.Add(Me.columnau_lname)
                Me.columnau_fname = New DataColumn("au_fname", GetType(System.String), Nothing, System.Data.MappingType.Element)
                Me.Columns.Add(Me.columnau_fname)
                Me.columnphone = New DataColumn("phone", GetType(System.String), Nothing, System.Data.MappingType.Element)
                Me.Columns.Add(Me.columnphone)
                Me.columnaddress = New DataColumn("address", GetType(System.String), Nothing, System.Data.MappingType.Element)
                Me.Columns.Add(Me.columnaddress)
                Me.columncity = New DataColumn("city", GetType(System.String), Nothing, System.Data.MappingType.Element)
                Me.Columns.Add(Me.columncity)
                Me.columnstate = New DataColumn("state", GetType(System.String), Nothing, System.Data.MappingType.Element)
                Me.Columns.Add(Me.columnstate)
                Me.columnzip = New DataColumn("zip", GetType(System.String), Nothing, System.Data.MappingType.Element)
                Me.Columns.Add(Me.columnzip)
                Me.columncontract = New DataColumn("contract", GetType(System.Boolean), Nothing, System.Data.MappingType.Element)
                Me.Columns.Add(Me.columncontract)
                Me.Constraints.Add(New UniqueConstraint("Constraint1", New DataColumn() {Me.columnau_id}, true))
                Me.columnau_id.AllowDBNull = false
                Me.columnau_id.Unique = true
                Me.columnau_lname.AllowDBNull = false
                Me.columnau_fname.AllowDBNull = false
                Me.columnphone.AllowDBNull = false
                Me.columncontract.AllowDBNull = false
            End Sub
            
            Public Function NewauthorsRow() As authorsRow
                Return CType(Me.NewRow,authorsRow)
            End Function
            
            Protected Overrides Function NewRowFromBuilder(ByVal builder As DataRowBuilder) As DataRow
                Return New authorsRow(builder)
            End Function
            
            Protected Overrides Function GetRowType() As System.Type
                Return GetType(authorsRow)
            End Function
            
            Protected Overrides Sub OnRowChanged(ByVal e As DataRowChangeEventArgs)
                MyBase.OnRowChanged(e)
                If (Not (Me.authorsRowChangedEvent) Is Nothing) Then
                    RaiseEvent authorsRowChanged(Me, New authorsRowChangeEvent(CType(e.Row,authorsRow), e.Action))
                End If
            End Sub
            
            Protected Overrides Sub OnRowChanging(ByVal e As DataRowChangeEventArgs)
                MyBase.OnRowChanging(e)
                If (Not (Me.authorsRowChangingEvent) Is Nothing) Then
                    RaiseEvent authorsRowChanging(Me, New authorsRowChangeEvent(CType(e.Row,authorsRow), e.Action))
                End If
            End Sub
            
            Protected Overrides Sub OnRowDeleted(ByVal e As DataRowChangeEventArgs)
                MyBase.OnRowDeleted(e)
                If (Not (Me.authorsRowDeletedEvent) Is Nothing) Then
                    RaiseEvent authorsRowDeleted(Me, New authorsRowChangeEvent(CType(e.Row,authorsRow), e.Action))
                End If
            End Sub
            
            Protected Overrides Sub OnRowDeleting(ByVal e As DataRowChangeEventArgs)
                MyBase.OnRowDeleting(e)
                If (Not (Me.authorsRowDeletingEvent) Is Nothing) Then
                    RaiseEvent authorsRowDeleting(Me, New authorsRowChangeEvent(CType(e.Row,authorsRow), e.Action))
                End If
            End Sub
            
            Public Sub RemoveauthorsRow(ByVal row As authorsRow)
                Me.Rows.Remove(row)
            End Sub
        End Class
        
        <System.Diagnostics.DebuggerStepThrough()>  _
        Public Class authorsRow
            Inherits DataRow
            
            Private tableauthors As authorsDataTable
            
            Friend Sub New(ByVal rb As DataRowBuilder)
                MyBase.New(rb)
                Me.tableauthors = CType(Me.Table,authorsDataTable)
            End Sub
            
            Public Property au_id As String
                Get
                    Return CType(Me(Me.tableauthors.au_idColumn),String)
                End Get
                Set
                    Me(Me.tableauthors.au_idColumn) = value
                End Set
            End Property
            
            Public Property au_lname As String
                Get
                    Return CType(Me(Me.tableauthors.au_lnameColumn),String)
                End Get
                Set
                    Me(Me.tableauthors.au_lnameColumn) = value
                End Set
            End Property
            
            Public Property au_fname As String
                Get
                    Return CType(Me(Me.tableauthors.au_fnameColumn),String)
                End Get
                Set
                    Me(Me.tableauthors.au_fnameColumn) = value
                End Set
            End Property
            
            Public Property phone As String
                Get
                    Return CType(Me(Me.tableauthors.phoneColumn),String)
                End Get
                Set
                    Me(Me.tableauthors.phoneColumn) = value
                End Set
            End Property
            
            Public Property address As String
                Get
                    Try 
                        Return CType(Me(Me.tableauthors.addressColumn),String)
                    Catch e As InvalidCastException
                        Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                    End Try
                End Get
                Set
                    Me(Me.tableauthors.addressColumn) = value
                End Set
            End Property
            
            Public Property city As String
                Get
                    Try 
                        Return CType(Me(Me.tableauthors.cityColumn),String)
                    Catch e As InvalidCastException
                        Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                    End Try
                End Get
                Set
                    Me(Me.tableauthors.cityColumn) = value
                End Set
            End Property
            
            Public Property state As String
                Get
                    Try 
                        Return CType(Me(Me.tableauthors.stateColumn),String)
                    Catch e As InvalidCastException
                        Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                    End Try
                End Get
                Set
                    Me(Me.tableauthors.stateColumn) = value
                End Set
            End Property
            
            Public Property zip As String
                Get
                    Try 
                        Return CType(Me(Me.tableauthors.zipColumn),String)
                    Catch e As InvalidCastException
                        Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                    End Try
                End Get
                Set
                    Me(Me.tableauthors.zipColumn) = value
                End Set
            End Property
            
            Public Property contract As Boolean
                Get
                    Return CType(Me(Me.tableauthors.contractColumn),Boolean)
                End Get
                Set
                    Me(Me.tableauthors.contractColumn) = value
                End Set
            End Property
            
            Public Function IsaddressNull() As Boolean
                Return Me.IsNull(Me.tableauthors.addressColumn)
            End Function
            
            Public Sub SetaddressNull()
                Me(Me.tableauthors.addressColumn) = System.Convert.DBNull
            End Sub
            
            Public Function IscityNull() As Boolean
                Return Me.IsNull(Me.tableauthors.cityColumn)
            End Function
            
            Public Sub SetcityNull()
                Me(Me.tableauthors.cityColumn) = System.Convert.DBNull
            End Sub
            
            Public Function IsstateNull() As Boolean
                Return Me.IsNull(Me.tableauthors.stateColumn)
            End Function
            
            Public Sub SetstateNull()
                Me(Me.tableauthors.stateColumn) = System.Convert.DBNull
            End Sub
            
            Public Function IszipNull() As Boolean
                Return Me.IsNull(Me.tableauthors.zipColumn)
            End Function
            
            Public Sub SetzipNull()
                Me(Me.tableauthors.zipColumn) = System.Convert.DBNull
            End Sub
        End Class
        
        <System.Diagnostics.DebuggerStepThrough()>  _
        Public Class authorsRowChangeEvent
            Inherits EventArgs
            
            Private eventRow As authorsRow
            
            Private eventAction As DataRowAction
            
            Public Sub New(ByVal row As authorsRow, ByVal action As DataRowAction)
                MyBase.New
                Me.eventRow = row
                Me.eventAction = action
            End Sub
            
            Public ReadOnly Property Row As authorsRow
                Get
                    Return Me.eventRow
                End Get
            End Property
            
            Public ReadOnly Property Action As DataRowAction
                Get
                    Return Me.eventAction
                End Get
            End Property
        End Class
    End Class
End Namespace
