Imports DotNetClient.localhost

Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents Button2 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.Button2 = New System.Windows.Forms.Button()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGrid1
        '
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(16, 40)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(264, 208)
        Me.DataGrid1.TabIndex = 0
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(16, 8)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(264, 23)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "GetAuthorsAsXml"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(296, 266)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button2, Me.DataGrid1})
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim svc As DataSetService = New DataSetService()
        Dim auset As GetAuthorsAsXmlResponseGetAuthorsAsXmlResult = svc.GetAuthorsAsXml()
        Dim au As AuthorType
        Dim row As Integer = 0
        Dim dt As DataTable = New DataTable("authors")
        dt.Columns.Add("id", GetType(System.String))
        dt.Columns.Add("au_first", GetType(System.String))
        dt.Columns.Add("au_last", GetType(System.String))

        For Each au In auset.AuthorSet
            Dim rowvals(2) As Object
            rowvals(0) = au.au_id
            rowvals(1) = au.au_fname
            rowvals(2) = au.au_lname
            dt.Rows.Add(rowvals)
        Next
        DataGrid1.DataSource = dt
    End Sub

End Class
