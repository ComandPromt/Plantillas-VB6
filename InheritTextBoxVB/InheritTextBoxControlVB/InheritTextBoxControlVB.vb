Public Class InheritTextBoxControlVB
    Inherits System.Windows.Forms.TextBox

    ' Check to see if string in text box appears to be a valid e-mail address.
    ' In this case, that means that it contains at least one @ sign and
    ' at least one period.

    Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)

        ' Pass call to base class

        MyBase.OnTextChanged(e)

        ' Perform our checking logic using inherited property Text, and
        ' set value of inherited property BackColor accordingly

        If (Me.Text.IndexOf("@") <> -1 And Me.Text.IndexOf(".") <> -1) Then
            Me.BackColor = System.Drawing.Color.LightGreen
        Else
            Me.BackColor = System.Drawing.Color.LightPink
        End If

    End Sub

End Class
