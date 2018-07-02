
Public Class BlinkingLabelControl
    Inherits System.Windows.Forms.Control


    ' Variables used internally

    Private myTimer As New System.Timers.Timer()
    Private m_BlinkInterval As Integer = 1
    Private m_BlinkoffColor As System.Drawing.Color = System.Drawing.Color.FromKnownColor(Drawing.KnownColor.Control)
    Private m_BlinkOnBrush As System.Drawing.SolidBrush
    Private m_BlinkOffBrush As System.Drawing.SolidBrush
    Private m_UseBlinkOnColor As Boolean = True


#Region " Component Designer generated code "

    Public Sub New()
        MyBase.New()

        ' Hook up handler function for timer, and set timer variables

        AddHandler myTimer.Elapsed, AddressOf OnTimerExpired
        myTimer.AutoReset = True
        myTimer.Interval = 1000
        myTimer.Enabled = True

        ' Create brushes used for drawing label text

        m_BlinkOnBrush = New System.Drawing.SolidBrush(Me.ForeColor)
        m_BlinkOffBrush = New System.Drawing.SolidBrush(Me.BackColor)

    End Sub

#End Region

    ' This function gets called when the control needs painting

    Protected Overrides Sub OnPaint(ByVal pe As System.Windows.Forms.PaintEventArgs)

        ' call the base class

        MyBase.OnPaint(pe)

        Dim BrushToUse As System.Drawing.Brush

        ' Choose the brush to use for the text color. If the blink cycle is
        ' currently on, or if we're in design mode (in which case we never
        ' want to blink, select the first brush. Otherwise select the second.

        If (m_UseBlinkOnColor = True Or Me.DesignMode = True) Then
            BrushToUse = m_BlinkOnBrush
        Else
            BrushToUse = m_BlinkOffBrush
        End If

        ' Draw the control's current Text property

        pe.Graphics.DrawString(Me.Text, Me.Font, BrushToUse, 0, 0)

    End Sub

    ' Control properties. This is the interval at which the blinking label
    ' changes its color from on to off or back again, in seconds.

    Public Property BlinkInterval() As Integer
        Get
            Return m_BlinkInterval
        End Get

        Set(ByVal Value As Integer)
            m_BlinkInterval = Value
            myTimer.Interval = m_BlinkInterval * 1000
        End Set
    End Property

    ' This property is the color in which the blinking label's text is shown
    ' during the Off part of the blink cycle. This allows you to make the blinking
    ' label appear to turn on and off, or to switch from one color to another.
    ' The control's ForeColor property is always used as the BlinkOn color.

    Public Property BlinkOffColor() As System.Drawing.Color
        Get
            Return m_BlinkoffColor
        End Get
        Set(ByVal Value As System.Drawing.Color)
            m_BlinkoffColor = Value
            m_BlinkOffBrush = New System.Drawing.SolidBrush(m_BlinkoffColor)
        End Set
    End Property

    ' When the ForeColor property, which we inherit from the base class, changes,
    ' then create a new brush with the new ForeColor to use in painting the label text.

    Private Sub CustomControl1_ForeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.ForeColorChanged
        m_BlinkOnBrush = New System.Drawing.SolidBrush(Me.ForeColor)
    End Sub


    ' Handler for our control's internal timer. 

    Private Sub OnTimerExpired(ByVal Source As Object, ByVal e As System.Timers.ElapsedEventArgs)

        ' Toggle the flag the tells the painting code whether to use the BlinkOnColor
        ' or the BlinkOff color

        If (m_UseBlinkOnColor = True) Then
            m_UseBlinkOnColor = False
        Else
            m_UseBlinkOnColor = True
        End If

        ' Invalidate the control to force a repaint.

        Me.Invalidate()

        ' Fire the blink event to the control's container, in case it cares.

        RaiseEvent BlinkStateChanged(m_UseBlinkOnColor)

    End Sub

    ' Declare an event that this control will fire to its container. This
    ' event is called Blink, and is fired when the internal timer expires
    ' and the blinking label changes color. The single parameter tells the
    ' event recipient whether the color is changing to BlinkOn (true) or to
    ' BlinkOff (false).

    Public Event BlinkStateChanged(ByVal UseBlinkOnColor As Boolean)

End Class
