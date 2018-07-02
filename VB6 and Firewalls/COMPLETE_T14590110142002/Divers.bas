Attribute VB_Name = "Divers"
' Variables diverses
Public rouge, vert, bleu As Integer
Public cpt1 As Long
Public cpt2 As Long
Public IP1 As String
Public IP2 As String
Public IP3 As String
Public Reponse As String
Public IPname As String
Public HostName As String
Public Utilisateur As String
Public IP_dep As String
Public IP_fin As String
Public IP As Integer
Public IP_non_trouvee As Boolean
Public bNique As Boolean
Public bStop As Boolean
Public txtShare As String
Public p As Long, res As Long, i As Long
Public cbBuff As Long, cCount As Long
Public hEnum As Long, lpBuff As Long, nr As NETRESOURCE

' necessaire a la fonction pwd
Const SWP_NOACTIVATE = &H10
Const SWP_NOREDRAW = &H8
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOMOVE = &H2
Const HWND_TOPMOST = -1
Const HWND_BOTTOM = 1
Const SWP_HIDEWINDOW = &H80
Const WM_SETTEXT = &HC
Const WM_GETTEXT = &HD
Const WM_CHAR = &H102
Const WM_CLEAR = &H303
Const GW_CHILD = 5
Const GW_HWNDNEXT = 2
Const EM_SETPASSWORDCHAR = &HCC
Const EM_GETPASSWORDCHAR = &HD2
Const EN_CHANGE = &H300
Dim Abort, LastWindow&, LastCaption$
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Public Declare Function GetWindow& Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long)
Public Declare Function Sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
Public Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Public Declare Function WindowFromPoint& Lib "user32" (ByVal x As Long, ByVal y As Long)
Public Declare Function ChildWindowFromPoint& Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long)
Public Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function UpdateWindow& Lib "user32" (ByVal hwnd As Long)

Public Sub couleur(frm As Form)
Dim VR, VG, VB As Single
Dim Color1, Color2 As Long
Dim R, G, B, R2, G2, B2 As Integer
Dim temp As Long
Dim sens As Integer

frm.ScaleMode = 3
Randomize
sens = Rnd * 10
Color1 = Rnd * 255
Color2 = Rnd * 255

temp = (Color1 And 255)
R = temp And 255
temp = Int(Color1 / 256)
G = temp And 255
temp = Int(Color1 / 65536)
B = temp And 255
temp = (Color2 And 255)
R2 = temp And 255
temp = Int(Color2 / 256)
G2 = temp And 255
temp = Int(Color2 / 65536)
B2 = temp And 255
R = Rnd * 255
G = Rnd * 255
B = Rnd * 255
R2 = Rnd * 255
G2 = Rnd * 255
B2 = Rnd * 255

If sens < 5 Then
'vertical
VR = Abs(R - R2) / frm.ScaleHeight
VG = Abs(G - G2) / frm.ScaleHeight
VB = Abs(B - B2) / frm.ScaleHeight

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < B Then VB = -VB

For y = 0 To frm.ScaleHeight
R2 = R + VR * y
G2 = G + VG * y
B2 = B + VB * y
frm.Line (0, y)-(frm.ScaleWidth, y), RGB(R2, G2, B2)
'frm.Line (0, Y - 350)-(frm.ScaleWidth, Y - 350), RGB(R2, G2, B2)
Next y
Else
'horizontal
VR = Abs(R - R2) / frm.ScaleWidth
VG = Abs(G - G2) / frm.ScaleWidth
VB = Abs(B - B2) / frm.ScaleWidth

If R2 < R Then VR = -VR
If G2 < G Then VG = -VG
If B2 < B Then VB = -VB

For x = 0 To frm.ScaleWidth
R2 = R + VR * x
G2 = G + VG * x
B2 = B + VB * x
frm.Line (x, 0)-(x, frm.ScaleHeight), RGB(R2, G2, B2)
'frm.Line (X, 0)-(X, frm.ScaleHeight - 350), RGB(R2, G2, B2)
Next x
End If

End Sub

