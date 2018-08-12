Attribute VB_Name = "ModDBCon"
'Designed and developed by Chris Hatton. if you want to reuse this code please email me.
'chris@hatton.com

Public cn As ADODB.Connection
Public MSMDB As String
Public Custom As Boolean
Public Function DBConnected() As Boolean

If Custom = False Then MSMDB = App.Path & "\" & "GALLERY.MDB"
    

On Error GoTo OpenDBError
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Open MSMDB
    DBConnected = True
Exit Function


OpenDBError:
DBConnected = False

MsgBox "Error Opening Database" & vbNewLine & Err.Description, vbCritical + vbOK
End




End Function
