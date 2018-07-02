Attribute VB_Name = "htmlmanager"
Global src, nme, lname As String

Sub Main()
makerepdir
Form1.Visible = True
End Sub

Public Function filterlinks()
'removes links and banners
While InStr(1, src, "<a href", 1) > 0
temp1 = InStr(1, src, "<a", 1)
temp2 = InStr(temp1, src, "</a>", 1)
src = Left(src, temp1 - 1) & Right(src, (Len(src) - (temp2 + 3)))
Wend
While InStr(1, src, "<iframe", 1) > 0
temp1 = InStr(1, src, "<ifram", 1)
temp2 = InStr(temp1, src, "</ifra", 1)
src = Left(src, temp1 - 1) + Right(src, Len(src) - temp2 - 8)
Wend

End Function
Public Function getname()
' Extracts country number
If InStr(1, src, "(#", 1) > 0 Then
flag1 = "(#"
flag2 = ")"
temp1 = InStr(1, src, flag1, 1)
temp2 = InStr(temp1, src, flag2, 1)
stp = temp2 - temp1
nme = Mid(src, temp1, stp + 1)
nme = nme & Date & ".htm"
Else
Countrynme = InputBox("What is the country Number For The Military Spy?", "Add Military Spy")
nme = "#" & Countrynme & "-" & "MS" & Left(Date, 5) & ".htm"
End If



End Function

Public Function savespy()
'um self-explanatory
Form1.cd.Filter = "Html|*.htm"
Form1.cd.InitDir = App.Path & "\Reports"
Form1.cd.FileName = nme
Form1.cd.ShowSave
If Len(Form1.cd.FileTitle) > 0 Then
Open Form1.cd.FileName For Output As #1
    Print #1, src
Close #1
lname = Form1.cd.FileName
showbrowser
End If

End Function
Public Function makerepdir()
'makes a Directory for The spy reports
On Error GoTo err:
MkDir (App.Path + "\reports")
err:

End Function



Public Function alloff()
'Turns all controls to vibible=false

Form1.addlbl.Visible = False
Form1.htmltxt.Visible = False
Form1.add.Visible = False
Form1.loaddir.Visible = False
Form1.loadfile.Visible = False
Form1.load.Visible = False
Form1.load.Enabled = False
Form1.clear.Enabled = False


End Function

Public Function showadd()
'shows controls to add an html
Form1.addlbl.Visible = True
Form1.htmltxt.Visible = True
Form1.add.Visible = True
Form1.clear.Enabled = False



End Function
Public Function showload()
'shows controls to load a html
Form1.loaddir.Visible = True
Form1.loadfile.Visible = True
Form1.load.Visible = True
Form1.loaddir.Path = App.Path + "\reports"
Form1.loadfile.Path = App.Path + "\reports"
End Function

Public Function showbrowser()
'shows spy report
If LCase(Right(lname, 3)) = "htm" Then
Form2.wb.Navigate (lname)
Form2.Visible = True
End If

End Function
Public Function editbody()
' replaces backgorund image with a black background
If InStr(1, src, "<body", 1) > 0 Then
    temp1 = InStr(1, src, "<body", 1)
    temp2 = InStr(temp1, src, ">", 1)
    src = Left(src, temp1) & "body bgcolor=black text=#fffffe>" & Right(src, (Len(src) - temp2))
End If
' Removes Your Status Heading (Turns,Food, Money, Nw)
If InStr(1, src, "<TABLE CELLSPACING=5>", 1) > 0 Then
    temp1 = InStr(1, src, "<TABLE CELLSPACING=5>", 1)
    temp2 = InStr(temp1, src, "</table>", 1)
src = Left(src, temp1 - 1) + Right(src, Len(src) - temp2 - 7)
End If
'removes everything after the spy
If InStr(1, src, "<TABLE WIDTH=400>", 1) > 0 Then
temp1 = InStr(1, src, "<TABLE WIDTH=400>", 1)
temp2 = InStr(temp1, src, "</form>", 1)
src = Left(src, temp1 - 1) + Right(src, Len(src) - temp2 + 1)
End If


End Function
