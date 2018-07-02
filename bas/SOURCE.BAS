Attribute VB_Name = "Draw"
Public Sub spixelset(x1, y1, x2, y2, Form As Form, style)
Form.ScaleMode = 3
Form.AutoRedraw = True
Select Case style
Case 0
Form.Line (x1, y1)-(x2, y2), 8421504
Form1.Line (x1, y1 + 1)-(x2, y2 + 1), 14803425

Case 1
Form.Line (x1, y1)-(x2, y2), 8421504
Form.Line (x1 + 1, y1)-(x2 + 1, y2), 14803425

Case 2
Form.Line (x1, y1)-(x2, y2), 14803425
Form.Line (x1 + 1, y1)-(x2 + 1, y2), 8421504
End Select
End Sub
Public Sub DrawLine(x1, y1, x2, y2, Form As Form, style)
Form.ScaleMode = 3
Select Case style
Case 0
Form.Line (x1, y1)-(x2, y2), 16777215
Form.Line (x1 + 1, y1)-(x2 + 1, y2), 12632256
Form.Line (x1 + 2, y1)-(x2 + 2, y2), 8421504
Form.Line (x1, y2)-(x2 + 2, y2), 8421504

Case 1
Form.Line (x1, y1)-(x2, y2), 16777215
Form.Line (x1, y1 + 1)-(x2, y2 + 1), 12632256
Form.Line (x1, y1 + 2)-(x2, y2 + 2), 8421504
End Select
End Sub

Public Sub High(Pic As PictureBox, x, y, Form As Form)
Form.ScaleMode = 3
Form.Line (x - 5, y - 5)-(Pic.Width + 5 + x, y - 5), 16777215
Form.Line (Pic.Width + 5 + x, y - 5)-(Pic.Width + x + 5, Pic.Height + 5 + y), 0
Form.Line (x - 5, Pic.Top + Pic.Height + 5)-(Pic.Width + 5 + x, Pic.Height + 5 + Pic.Top), 0
Form.Line (Pic.Left - 5, Pic.Top - 5)-(Pic.Left - 5, Pic.Top + Pic.Height + 5), 16777215
End Sub

Public Sub Eras(Pic As PictureBox, x, y, Form As Form)
Form.ScaleMode = 3
Form.Line (x - 5, y - 5)-(Pic.Width + 5 + x, y - 5), Form.BackColor
Form.Line (Pic.Width + 5 + x, y - 5)-(Pic.Width + x + 5, Pic.Height + 5 + y), Form.BackColor
Form.Line (x - 5, Pic.Top + Pic.Height + 5)-(Pic.Width + 5 + x, Pic.Height + 5 + Pic.Top), Form.BackColor
Form.Line (Pic.Left - 5, Pic.Top - 5)-(Pic.Left - 5, Pic.Top + Pic.Height + 5), Form.BackColor
End Sub
