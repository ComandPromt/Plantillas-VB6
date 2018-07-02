Attribute VB_Name = "Module1"
Public Sub GridMultiSelect _
(Button As Integer, Shift As Integer, X As Single, Y As Single, ThisGrid As DBGrid, ThisRS As Recordset)
    
'Sub written by Gary Thibault
'If you have any improvements or commments then please contact me
'at gwtbolt@frontiernet.net

'Create a mouse down event for the grid you want to apply this sub to.
'Your mouse down event should look like this:
'Private Sub DBgrid_MouseDown _
'(Button As Integer, Shift As Integer, X As Single, Y As Single)
'DBgrid should be changed to the name of your DBgrid

'Next, call the sub. The call should look like this
'Call GridMultiSelect(Button, Shift, X, Y, DBgrid, Datcontrol.Recordset)
'All you have to do to the above call is change DBgrid to the name of your grid and
'change DatControl to the name of your Datcontrol. Everthing else stays the same
    
    'X = Col position and Y = Row Position
Dim LeftButtonClicked As Boolean
Dim ShiftButtonPressed As Boolean
Dim OverRowSelector As Boolean
Dim SelFirst As Integer
Dim SelLast As Integer
Dim BeginRow As Integer
Dim EndRow As Integer
Dim Bmk As Variant
Dim Row As Integer
Dim SelLastGridRow As Single

    'returns true if Left mouse button is clicked
LeftButtonClicked = (Button = vbLeftButton)
    'returns true if shift button is held down
ShiftButtonPressed = (Shift = vbShiftMask)
    'returns true only if a row selecter is clicked,
    'remove this if you want the user to be able to click
    'anywhere in the row to select that row.
OverRowSelector = (ThisGrid.ColContaining(X) = -1)
    'Clicked over row selector with shift button pressed and
    'one row previously selected as your starting row.
    'Also makes sure user did not select previously selected row
If LeftButtonClicked And OverRowSelector And ShiftButtonPressed _
And ThisGrid.SelBookmarks.Count = 1 And _
ThisGrid.Row <> ThisGrid.RowContaining(Y) Then

    'get the relative record number for the first record
SelFirst = ThisRS.AbsolutePosition
    'get the bookmark for the next item selected in the grid
Bmk = ThisGrid.RowBookmark(ThisGrid.RowContaining(Y))
    'move the current record to the next item selected because
    'the current record did not change while shift key was pressed
ThisRS.Bookmark = Bmk
    'get the relative record number for the last record
SelLast = ThisRS.AbsolutePosition
    'record the *grids* row number for the last item selected
SelLastGridRow = ThisGrid.Row

 
 On Error GoTo ErrorHandler
        'make sure that we are looping from low to high
    If SelFirst < SelLast Then
        BeginRow = SelFirst
        EndRow = SelLast
    Else
        BeginRow = SelLast
        EndRow = SelFirst
    End If
        
    
        'add all the bookmarks to the selbookmark collection
    For Row = BeginRow To EndRow
        ThisRS.AbsolutePosition = Row
        Bmk = ThisRS.Bookmark
        ThisGrid.SelBookmarks.Add Bmk
    Next Row
        'Return display to original viewing position while
        'moving selector to record clicked with shift key
    
    
    If ThisGrid.RowBookmark(ThisGrid.Row) = ThisGrid.FirstRow Then
        ThisRS.AbsolutePosition = SelLast
            'converts the positive into a negative
        SelLastGridRow = (SelLastGridRow - (SelLastGridRow * 2))
        ThisGrid.Scroll 0, SelLastGridRow
    End If
    
End If
Exit Sub
'why crash?
ErrorHandler:
MsgBox Err.Number & " " & Err.Description
End Sub


