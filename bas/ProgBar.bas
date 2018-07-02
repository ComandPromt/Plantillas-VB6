Attribute VB_Name = "Module1"
Option Explicit
Global gbCancel As Boolean

Sub Main()
    Dim lX As Long, lY As Long, lR As Long
    
    For lX = 1 To 50                                                        'this is just a loop to take up time
        For lY = 1 To 10
            ProgBar "Sample #1", lY, 10, 7, lX, 50                          'show the progress bar
            DoEvents                                                        'make sure the computer does not "hog" resources
            If gbCancel Then End                                            'if the user pressed Cancel, then quit
        Next lY
    Next lX
    lR = MsgBox("Ready for the next sample?", vbOKCancel, "You Choose:")    'Ask if the user wants to see #2
    Unload frmProgress                                                      'unload the form
    Set frmProgress = Nothing                                               'free up the memory the frmProgress took up
    If lR = 2 Then End                                                      'if user presses Cancel on the msgbox, then quit
    lX = 0: lY = 0                                                          'reset the variables
    For lX = 1 To 50                                                        'loop again
        For lY = 1 To 10
            ProgBar "Sample #2", lX, 50                                     'show the progress bar
            DoEvents                                                        'make sure the computer does not "hog" resources
            If gbCancel Then End                                            'if the user pressed Cancel, then quit
        Next lY
    Next lX
    Unload frmProgress                                                      'unload the form
    Set frmProgress = Nothing                                               'free up the memory the frmProgress took up
    MsgBox "The End!"                                                       'The End
    End                                                                     'End
End Sub

Sub ProgBar(sCap As String, lLevel As Long, lMaxLevel As Long, Optional iColor As Integer, Optional lTotal As Long, Optional lTotLevel As Long, Optional iColorTot As Integer)
'I found this code in a sample program I downloaded.  I then converted it (massively) into what it is now.
'       Eric D. Burdo
' Description:  A horizontal progress bar.
' Parameters:
'sCap       = The Caption for the Progress box.
'lLevel     = the current level for the progress bar to be
'lMaxLevel  = the maximum the bar can go to
'** The following are all Optional parameters.
'iColor     = the color (using QBColor) that you want the progress bar.
'bTotal     = True if you want a Total progress bar, False (default) if you do not
'lTotal     = the current level for the Total progress bar to be
'lTotLevel  = the maximum the Total bar can go to
'iColorTot     = the color (using QBColor) that you want the Total progress bar.
'Sample:    ProgBar "SaberZip Progress", CLng(lCur), CLng(lCount), 5, lCount, mlTotal, 8
'Remember:  The totals need to be Long variables.  If your variable is not Long that you are passing,
'           then convert it using the Clng(Integer)

    Dim zSingleUnit As Single
    Dim zSingleUnitTot As Single
    Dim f As Form
    
    Set f = frmProgress
    'Set the defaults and beginning variables
    f.Caption = sCap                                                                    'Set the caption
    If iColor = 0 Or iColor = 7 Then iColor = 5                                         'if a color is not specified, or it is 7
    If iColorTot = 0 Or iColor = 7 Then iColorTot = 8                                   'then use different colors.  7 is the same color as the background
    On Error Resume Next
    zSingleUnit = f.picGauge.ScaleWidth / lMaxLevel                                     'set the single until values
    zSingleUnitTot = f.picGaugeTotal.ScaleWidth / lTotLevel                             'set the single until values
    
    ' Draw bar.
    With f.picGauge                                                                     'this is the regular progress bar
        f.picGauge.Line (.ScaleLeft, .ScaleTop)-Step(lLevel * zSingleUnit, .ScaleHeight), QBColor(iColor), BF 'RGB(0, 0, 255)
        If lLevel >= lMaxLevel Then
            f.picGauge.Line (.ScaleLeft, .ScaleTop)-Step(.ScaleWidth, .ScaleHeight), QBColor(iColor), BF
        End If
        f.picGauge.Line (lLevel * zSingleUnit, .ScaleTop)-Step(.ScaleWidth, .ScaleHeight), .BackColor, BF
    End With
    f.lblPercentVal.Caption = LTrim$(Str$(Fix(100 * (lLevel / lMaxLevel)))) + "%"
    f.Label1.Caption = LTrim$(Str$(Fix(100 * (lLevel / lMaxLevel)))) + "%"
    If Not lTotal = 0 Then                                                              'do we need a total progress bar?
        f.fraPercentTotal.Visible = True
        f.cmdCancel.Top = f.fraPercent.Top + f.fraPercentTotal.Top + 1200               'Set button vertical position
        f.Height = f.cmdCancel.Top + 1080
        f.lblPercentValTotal.Caption = LTrim$(Str$(Fix(100 * (lTotal / lTotLevel)))) + "%"
        With f.picGaugeTotal
            f.picGaugeTotal.Line (.ScaleLeft, .ScaleTop)-Step(lTotal * zSingleUnitTot, .ScaleHeight), QBColor(iColorTot), BF 'RGB(0, 0, 255)
            If lTotal >= lTotLevel Then
                f.picGaugeTotal.Line (.ScaleLeft, .ScaleTop)-Step(.ScaleWidth, .ScaleHeight), QBColor(iColorTot), BF
            End If
            f.picGaugeTotal.Line (lTotal * zSingleUnitTot, .ScaleTop)-Step(.ScaleWidth, .ScaleHeight), .BackColor, BF
        End With
    Else                                                                                'we only want the single bar, so set the height and such.
        frmProgress.fraPercentTotal.Visible = False
        f.cmdCancel.Top = f.fraPercent.Top + 1200                                       'Set button vertical position
            If f.cmdCancel.Top < 970 Then f.cmdCancel.Top = 970                         'Use minimum vertical
        f.Height = f.cmdCancel.Top + 1080                                               'Set window height
    End If
    f.Show                                                                              'Show the form
    f.Refresh                                                                           'refresh the screen so it shows properly if running fast
    DoEvents
End Sub
