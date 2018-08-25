Attribute VB_Name = "Module1"
'**************************************
'Windows API/Global Declarations for :*M
'     ove a Form without even having a Title B
'     ar!!!*
'**************************************
'Type the following in the Module/Bas!!
'     NOT IN THE FORM!! (it wont work!)


Declare Sub ReleaseCapture Lib "user32" ()


Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long


Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

