Public Class UserControl1
    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'UserControl1 overrides dispose to clean up the component list.
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "User ID:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Password:"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(40, 144)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "OK"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(144, 144)
        Me.Button2.Name = "Button2"
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Cancel"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(24, 32)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(176, 20)
        Me.TextBox1.TabIndex = 4
        Me.TextBox1.Text = ""
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(24, 88)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(176, 20)
        Me.TextBox2.TabIndex = 5
        Me.TextBox2.Text = ""
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.DataMember = Nothing
        '
        'UserControl1
        '
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TextBox2, Me.TextBox1, Me.Button2, Me.Button1, Me.Label2, Me.Label1})
        Me.Name = "UserControl1"
        Me.Size = New System.Drawing.Size(256, 208)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim bFieldsValid As Boolean = True

        ' Check to make sure that required fields are filled in.
        ' Set error provider control to signal errors to the user if they're not.

        If (TextBox1.Text.Length = 0) Then
            ErrorProvider1.SetError(TextBox1, "A UserID is required")
            bFieldsValid = False
        Else
            ErrorProvider1.SetError(TextBox1, "")
        End If

        If (TextBox2.Text.Length = 0) Then
            ErrorProvider1.SetError(TextBox2, "A Password is required")
            bFieldsValid = False
        Else
            ErrorProvider1.SetError(TextBox2, "")
        End If

        ' Fire event to container if they are. 

        If (bFieldsValid = True) Then
            RaiseEvent OkClicked(TextBox1.Text, TextBox2.Text)
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        RaiseEvent CancelClicked()
    End Sub

    ' Custom property, the background color used by both 
    ' constituent text boxes

    Private m_BothTextBoxesBackColor As System.Drawing.Color = Color.FromKnownColor(KnownColor.Window)

    Public Property BothTextBoxesBackColor() As System.Drawing.Color
        Get
            Return m_BothTextBoxesBackColor
        End Get
        Set(ByVal Value As System.Drawing.Color)
            m_BothTextBoxesBackColor = Value
            TextBox1.BackColor = m_BothTextBoxesBackColor
            TextBox2.BackColor = m_BothTextBoxesBackColor
        End Set
    End Property

    ' Events that this control fires to its container when the user clicks the
    ' OK button or the Cancel button.

    Public Event OkClicked(ByVal UserID As String, ByVal Password As String)
    Public Event CancelClicked()

    Private m_Pen As New System.Drawing.Pen(Color.Black, 3)

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        'Me.OnPaint(e)
        e.Graphics.DrawRectangle(m_Pen, 0, 0, Me.Bounds.Width - 1, Me.Bounds.Height - 1)
    End Sub
End Class
