Attribute VB_Name = "Picture"
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Sub ReleaseCapture Lib "user32" ()

Sub AddPic2RTB()
SetParent frmInsert.Picture1.hWnd, frmMain.Text1.hWnd
SetParent frmMain.Picture1.hWnd, frmMain.Text1.hWnd

End Sub


Sub RemovePic2RTB()
SetParent frmInsert.Picture1.hWnd, frmInsert.hWnd
End Sub
