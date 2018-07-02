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
    Friend WithEvents UserControl11 As UserControlLibraryVB.UserControl1
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.UserControl11 = New UserControlLibraryVB.UserControl1()
        Me.SuspendLayout()
        '
        'UserControl11
        '
        Me.UserControl11.BothTextBoxesBackColor = System.Drawing.Color.PaleGoldenrod
        Me.UserControl11.Location = New System.Drawing.Point(32, 16)
        Me.UserControl11.Name = "UserControl11"
        Me.UserControl11.Size = New System.Drawing.Size(256, 208)
        Me.UserControl11.TabIndex = 0
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(336, 245)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.UserControl11})
        Me.Name = "Form1"
        Me.Text = "Rolling Thunder UserControl Client VB"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub UserControl11_CancelClicked() Handles UserControl11.CancelClicked
        MessageBox.Show("User clicked Cancel button")
    End Sub

    Private Sub UserControl11_OkClicked(ByVal UserID As String, ByVal Password As String) Handles UserControl11.OkClicked
        MessageBox.Show("User clicked OK button UserID = " + UserID + ", Password = " + Password)
    End Sub
End Class
