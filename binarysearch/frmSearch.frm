VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Binary Search"
   ClientHeight    =   2520
   ClientLeft      =   3765
   ClientTop       =   2385
   ClientWidth     =   6975
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   2040
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000000C0&
      Caption         =   "Select a Record to Search For"
      ForeColor       =   &H0000FFFF&
      Height          =   2295
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton cmdSearchArray 
         Caption         =   "Search Array"
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearchFile 
         Caption         =   "Search File"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox lstCustID 
         BackColor       =   &H000000FF&
         ForeColor       =   &H0000FFFF&
         Height          =   1425
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtID 
         BackColor       =   &H000000FF&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Caption         =   "Search Results"
      ForeColor       =   &H0000FFFF&
      Height          =   1815
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtCustID 
         BackColor       =   &H000000FF&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtCompanyName 
         BackColor       =   &H000000FF&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtContact 
         BackColor       =   &H000000FF&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtPhone 
         BackColor       =   &H000000FF&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblCustomerID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID:"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   915
      End
      Begin VB.Label lblCompanyName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name:"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblContact 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contact:"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label lblPhone 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   1440
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' Customer record user defined type.
'
Private Type tCust
    CustID   As String * 5
    CompName As String * 40
    Contact  As String * 30
    Phone    As String * 24
End Type

'
' Array of customers
'
Private CustArray() As tCust
Private CustRec     As tCust

Private RecCt  As Long
Private Sub cmdQuit_Click()

    Unload Me
    
End Sub
Private Sub Form_Load()
Dim GetCust As tCust
Dim l       As Long

    '
    ' Load an array with data from the file and
    ' load the listbox with the ID from each record.
    '
    Open "CustList.dat" For Random As 1 Len = Len(GetCust)
    RecCt = LOF(1) / Len(GetCust)
    ReDim CustArray(1 To RecCt)
    
    For l = 1 To RecCt
        Get 1, l, CustArray(l)
        lstCustID.AddItem CustArray(l).CustID
    Next
    Close 1

End Sub
Private Sub pShowArrayRecord(lngRecord As Long)
        
    '
    ' Display a record from the array in the textboxes.
    '
    txtCustID = CustArray(lngRecord).CustID
    txtCompanyName = CustArray(lngRecord).CompName
    txtContact = CustArray(lngRecord).Contact
    txtPhone = CustArray(lngRecord).Phone
    
End Sub
Private Sub pShowFileRecord(lngRecord As Long)
    '
    ' Display a record from the file in the textboxes.
    '
    txtCustID = CustRec.CustID
    txtCompanyName = CustRec.CompName
    txtContact = CustRec.Contact
    txtPhone = CustRec.Phone
    
End Sub
Private Sub ClearRecord()

    '
    ' Clear the text boxes
    '
    txtCustID = ""
    txtCompanyName = ""
    txtContact = ""
    txtPhone = ""
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Erase CustArray
    Set frmSearch = Nothing
    
End Sub
Private Sub lstCustID_Click()
    '
    ' Display the selected listbox item in the textbox.
    '
    txtID = lstCustID.Text
    
End Sub
Private Sub cmdSearchArray_Click()
Dim lngMatch As Long
           
    '
    ' Search the array for a record.
    '
    lngMatch = fSearchArray(txtID)
    '
    ' If found, display the record.
    '
    If lngMatch Then
        Call pShowArrayRecord(lngMatch)
    Else
        Call ClearRecord
    End If
    
End Sub
Private Sub cmdSearchFile_Click()
Dim lngMatch As Long
    
    '
    ' Search the file for a record.
    '
    lngMatch = fSearchFile(txtID)
    '
    ' If found, display the record.
    '
    If lngMatch Then
        Call pShowFileRecord(lngMatch)
    Else
        Call ClearRecord
    End If

End Sub
Private Function fSearchArray(strSearchItem As String) As Long
Dim lngFirst    As Long
Dim lngLast     As Long
Dim lngMiddle   As Long
Dim lngLastPass As Long
Dim strItem     As String
Dim strValue    As String
Dim blnDone     As Boolean
    '
    ' Search an array for an item using a binary search.
    ' The search is not case sensitive.
    ' Returned is the index of the matching array element.
    '

    '
    ' Initialize the pointers to the first
    ' and last records.
    '
    lngFirst = 1
    lngLast = UBound(CustArray)
    strItem = UCase$(Trim$(strSearchItem))
    '
    ' If only one record, see if it is the desired one.
    '
    If lngLast = 1 Then
        If strItem = UCase$(CustArray(1).CustID) Then
            fSearchArray = 1
        Else
            fSearchArray = 0
        End If
        Exit Function
    End If
    '
    ' Set the pointer to the middle record.
    '
    lngMiddle = ((lngLast - lngFirst) + 1) \ 2

    '
    ' Apply the binary search criteria until the
    ' item is found or the list is exhausted.
    '
    Do Until blnDone
        strValue = UCase$(CustArray(lngMiddle).CustID)
        
        If strItem = strValue Then
            '
            ' Found it.
            '
            fSearchArray = lngMiddle
            blnDone = True
            Exit Do
        ElseIf strItem < strValue Then
            '
            ' Direction = down
            ' Remove the second half of the list.
            '
            lngLast = lngMiddle
            lngMiddle = lngMiddle - ((lngLast - lngFirst) + 1) \ 2
        ElseIf strItem > strValue Then
            '
            ' Direction = Up
            ' Remove the first half of the list.
            '
            lngFirst = lngMiddle
            lngMiddle = lngMiddle + ((lngLast - lngFirst) + 1) \ 2
        End If
        
        '
        ' See if list is still divisible.
        '
        If (lngMiddle = lngFirst) Or (lngMiddle = lngLast) Then
            lngLastPass = lngLastPass + 1
            If lngLastPass = 2 Then
                lngLastPass = 0
                fSearchArray = 0
                blnDone = True
            End If
        End If
    Loop
    
End Function
Private Function fSearchFile(strSearchItem As String) As Long
Dim lngFirst    As Long
Dim lngLast     As Long
Dim lngMiddle   As Long
Dim lngLastPass As Long
Dim strItem     As String
Dim strValue    As String
Dim blnDone     As Boolean


    Open "CustList.dat" For Random As 1 Len = Len(CustRec)
    RecCt = LOF(1) / Len(CustRec)
    '
    ' Search a file for an item using a binary search.
    ' The search is not case sensitive.
    ' Returned is the index of the matching file element.
    '

    '
    ' Initialize the pointers to the first
    ' and last records.
    '
    lngFirst = 1
    lngLast = RecCt
    strItem = UCase$(Trim$(strSearchItem))
    
    '
    ' If only one record, see if it is the desired one.
    '
    If lngLast = 1 Then
        Get 1, 1, CustRec
        If strItem = UCase$(CustRec.CustID) Then
            fSearchFile = 1
        Else
            fSearchFile = 0
        End If
        Close 1
        Exit Function
    End If
    '
    ' Set the pointer to the middle record.
    '
    lngMiddle = ((lngLast - lngFirst) + 1) \ 2

    '
    ' Apply the binary search criteria until the
    ' item is found or the file is exhausted.
    '
    Do Until blnDone
        '
        ' Read a record from the file.
        '
        Get 1, lngMiddle, CustRec
        
        strValue = UCase$(CustRec.CustID)
        
        If strItem = strValue Then
            '
            ' Found it.
            '
            fSearchFile = lngMiddle
            blnDone = True
            Exit Do
        ElseIf strItem < strValue Then
            '
            ' Direction = down
            ' Remove the second half of the file.
            '
            lngLast = lngMiddle
            lngMiddle = lngMiddle - ((lngLast - lngFirst) + 1) \ 2
        ElseIf strItem > strValue Then
            '
            ' Direction = Up
            ' Remove the first half of the file.
            '
            lngFirst = lngMiddle
            lngMiddle = lngMiddle + ((lngLast - lngFirst) + 1) \ 2
        End If
        
        '
        ' See if file is still divisible.
        '
        If (lngMiddle = lngFirst) Or (lngMiddle = lngLast) Then
            lngLastPass = lngLastPass + 1
            If lngLastPass = 2 Then
                lngLastPass = 0
                fSearchFile = 0
                blnDone = True
            End If
        End If
    Loop
    
    Close 1
    
End Function
Private Sub txtID_KeyPress(KeyAscii As Integer)
    '
    ' Convert to upper case.
    '
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    
End Sub


