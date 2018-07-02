Attribute VB_Name = "Effects"
Public Sub WipeRight(Lt%, Tp%, frm As Form)
    Dim s, Wx, Hx, i
    s = 50 'number of steps to use in the wipe
    Wx = frm.Width / s 'size of vertical steps
    Hx = frm.Height / s 'size of horizontal steps
    ' top and left are static
    ' while the width gradually shrinks
    For i = 1 To s - 1
        frm.Move Lt%, Tp%, frm.Width - Wx
    Next
End Sub
Public Sub Explode(frm As Form)
 '<<==Crap, needs debugging==>>
  ' frm.Width = 0
  ' frm.Height = 0
  ' frm.Show
   For x = 0 To 5000 Step 250
      frm.Width = x
      frm.Height = x
      With frm
         .Left = (Screen.Width - .Width) / 2
         .Top = (Screen.Height - .Height) / 2
      End With
   Next x
End Sub

