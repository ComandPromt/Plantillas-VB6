Attribute VB_Name = "Module1"
'***************************************************************
' Author: Andrew Jackson
' Date:   16/03/01
'
'***************************************************************

Public Sub Add2Tree(strBranch As String, tvTree As TreeView, strParent As String, intParentImage As Integer, intChildImage As Integer)

'Initialise Variables
Dim arrInfo() As String, strRelative As String, strKey As String

'Set Character that will split up String
Const ChrSplit = "\"

'Add backslash at start, this will create an empty cell in the lbound of array
strBranch = ChrSplit & UCase$(strBranch)

'Seperates the directory into individualy parts and puts them into an Array
Call GetWords(arrInfo(), strBranch, ChrSplit)

'Adds a parent to the treeview
'Checks the treeview for node key, the node is added if the Key is unique
If FindNode(strParent & arrInfo(0) & ChrSplit, tvTree) = False Then
    tvTree.Nodes.Add , , strParent & arrInfo(0) & ChrSplit, strParent, intParentImage
End If

'Loops though array adding nodes if they don't exist
For InfoCount = 1 To UBound(arrInfo)
    'Create a Relative
    strRelative = strRelative & arrInfo(InfoCount - 1)
    'If Backlash doesn't exist at the end, add it
    If Not Right$(strRelative, 1) = ChrSplit Then strRelative = strRelative & ChrSplit
    
    'Create a Key for the node
    strKey = strRelative & arrInfo(InfoCount)
    'If Backslash doesn't exist at the end, add it
    If Not Right$(strKey, 1) = ChrSplit Then strKey = strKey & ChrSplit

    'Adds child nodes to to treeview
    'Checks the treeview for node key, the node is added if the Key is unique
    If FindNode(strParent & strKey, tvTree) = False Then
        tvTree.Nodes.Add strParent & strRelative, tvwChild, strParent & strKey, arrInfo(InfoCount), intChildImage
    End If
Next InfoCount

'Remove Array from memory
Erase arrInfo()

End Sub

Private Function FindNode(NodeSearch As String, tvTree As TreeView) As Boolean

'Search though Treeview for existing node keys
For NodeCount = 1 To tvTree.Nodes.Count
    'Compares String with node key
    If NodeSearch = tvTree.Nodes(NodeCount).Key Then
        'If Node key is the same as String then key is not Unique and the
        'function will return true for found
        FindNode = True
        'Exit NodeCount
        Exit For
    End If
Next NodeCount

End Function

Private Sub GetWords(WordStore() As String, strSource As String, strSpltWth As String)

'Initialise Variables
Dim intLstPos As Integer, intSize As Integer

'Sets Start of String
intLstPos = 1

'Search for the specified characters within the string
'Loops though lenth of string
For Counter = 1 To Len(strSource)
    'Resizes array to the number of the specifed characters found
    ReDim Preserve WordStore(intSize)
    
    'Compares currant part or string with specified characters
    If Mid$(strSource, Counter, Len(strSpltWth)) = strSpltWth Then
        'Adds Section of string to ubound cell of array
        WordStore(intSize) = Trim$(Mid$(strSource, intLstPos, Counter - intLstPos))
        'Stores the new value that will be the next text Start
        intLstPos = Counter + Len(strSpltWth)
        'Adds one to array size variable
        intSize = intSize + 1
    End If
Next Counter

End Sub
