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
    Friend WithEvents InheritTextBoxControlVB1 As InheritTextBoxControlVB.InheritTextBoxControlVB
    Friend WithEvents Label1 As System.Windows.Forms.Label

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.InheritTextBoxControlVB1 = New InheritTextBoxControlVB.InheritTextBoxControlVB()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'InheritTextBoxControlVB1
        '
        Me.InheritTextBoxControlVB1.BackColor = System.Drawing.Color.LightPink
        Me.InheritTextBoxControlVB1.Location = New System.Drawing.Point(24, 40)
        Me.InheritTextBoxControlVB1.Name = "InheritTextBoxControlVB1"
        Me.InheritTextBoxControlVB1.Size = New System.Drawing.Size(200, 20)
        Me.InheritTextBoxControlVB1.TabIndex = 0
        Me.InheritTextBoxControlVB1.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "E-mail address:"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(304, 125)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label1, Me.InheritTextBoxControlVB1})
        Me.Name = "Form1"
        Me.Text = "Rolling Thunder InheritTextBox Client VB"
        Me.ResumeLayout(False)

    End Sub

#End Region

End Class
