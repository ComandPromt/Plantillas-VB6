Attribute VB_Name = "Module1"
Option Explicit
Public fMainForm As frmMain
   Private Type Rect
      left As Long
      Top As Long
      Right As Long
      Bottom As Long
   End Type

   Private Type CharRange
     cpMin As Long
     cpMax As Long
   End Type

   Private Type FormatRange
     hdc As Long
     hdcTarget As Long
     rc As Rect
     rcPage As Rect
     chrg As CharRange
   End Type

   Private Const WM_USER As Long = &H400
   Private Const EM_FORMATRANGE As Long = WM_USER + 57
   Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
   Private Const PHYSICALOFFSETX As Long = 112
   Private Const PHYSICALOFFSETY As Long = 113

   Private Declare Function GetDeviceCaps Lib "gdi32" ( _
      ByVal hdc As Long, ByVal nIndex As Long) As Long
   Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
      (ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, _
      lp As Any) As Long
   Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
      (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
      ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
      
      
  




 Public Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, _
      TopMarginHeight, RightMarginWidth, BottomMarginHeight)
      Dim LeftOffset As Long, TopOffset As Long
      Dim LeftMargin As Long, TopMargin As Long
      Dim RightMargin As Long, BottomMargin As Long
      Dim fr As FormatRange
      Dim rcDrawTo As Rect
      Dim rcPage As Rect
      Dim TextLength As Long
      Dim NextCharPosition As Long
      Dim r As Long

     
      Printer.Print space(1)
      Printer.ScaleMode = vbTwips

      
      LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
         PHYSICALOFFSETX), vbPixels, vbTwips)
      TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
         PHYSICALOFFSETY), vbPixels, vbTwips)

    
      LeftMargin = LeftMarginWidth - LeftOffset
      TopMargin = TopMarginHeight - TopOffset
      RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
      BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset

    
      rcPage.left = 0
      rcPage.Top = 0
      rcPage.Right = Printer.ScaleWidth
      rcPage.Bottom = Printer.ScaleHeight

      rcDrawTo.left = LeftMargin
      rcDrawTo.Top = TopMargin
      rcDrawTo.Right = RightMargin
      rcDrawTo.Bottom = BottomMargin

     
      fr.hdc = Printer.hdc
      fr.hdcTarget = Printer.hdc
      fr.rc = rcDrawTo
      fr.rcPage = rcPage
      fr.chrg.cpMin = 0
      fr.chrg.cpMax = -1

     
      TextLength = Len(RTF.Text)

    
      Do
        
         NextCharPosition = SendMessage(RTF.hWnd, EM_FORMATRANGE, True, fr)
         If NextCharPosition >= TextLength Then Exit Do
         fr.chrg.cpMin = NextCharPosition
         Printer.NewPage
         Printer.Print space(1)
         fr.hdc = Printer.hdc
         fr.hdcTarget = Printer.hdc
      Loop

    
      Printer.EndDoc

    
      r = SendMessage(RTF.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))
   End Sub
   
    

Public Function FileExists(strFile As String) As String


    On Error Resume Next
   
    FileExists = Dir(strFile, vbHidden) <> ""
    
End Function

