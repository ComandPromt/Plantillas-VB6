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
    Friend WithEvents BlinkingLabelControl1 As BlinkingLabelControlVB.BlinkingLabelControl

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.BlinkingLabelControl1 = New BlinkingLabelControlVB.BlinkingLabelControl()
        Me.SuspendLayout()
        '
        'BlinkingLabelControl1
        '
        Me.BlinkingLabelControl1.BlinkInterval = 1
        Me.BlinkingLabelControl1.BlinkOffColor = System.Drawing.SystemColors.Control
        Me.BlinkingLabelControl1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BlinkingLabelControl1.Location = New System.Drawing.Point(32, 24)
        Me.BlinkingLabelControl1.Name = "BlinkingLabelControl1"
        Me.BlinkingLabelControl1.Size = New System.Drawing.Size(192, 40)
        Me.BlinkingLabelControl1.TabIndex = 0
        Me.BlinkingLabelControl1.Text = "BlinkingLabelControl1"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(336, 149)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.BlinkingLabelControl1})
        Me.Name = "Form1"
        Me.Text = "Rolling Thunder Blinking Label Client VB"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub BlinkingLabelControl1_BlinkStateChanged(ByVal UseBlinkOnColor As Boolean) Handles BlinkingLabelControl1.BlinkStateChanged
        Beep()
    End Sub
End Class
