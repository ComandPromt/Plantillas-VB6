Attribute VB_Name = "projectMaintenanceCode"
Option Explicit

Public Function fillUserCombo(cboTmp As ComboBox)
  Dim rstUsers As New ADODB.Recordset
  Dim intPosition As Integer
  cboTmp.clear
  intPosition = 0
  Set rstUsers = returnRS(cmdSelectUsers)
  If rstUsers.EOF <> True Then
    rstUsers.MoveFirst
    While rstUsers.EOF <> True
      cboTmp.AddItem rstUsers![strFirstName] & " " & rstUsers![strLastName], intPosition
      cboTmp.ItemData(intPosition) = rstUsers![lngUserID]
      intPosition = intPosition + 1
      rstUsers.MoveNext
    Wend
  End If
End Function

Public Function returnUserID(cboTmp As ComboBox) As Long
  If cboTmp.ListIndex <> -1 Then
    returnUserID = cboTmp.ItemData(cboTmp.ListIndex)
  Else
    returnUserID = 0
  End If
End Function

Public Function matchUserID(cboTmp As ComboBox, lngUserID As Long) As Integer
  Dim intItems As Integer
  For intItems = 0 To cboTmp.ListCount - 1
    If cboTmp.ItemData(intItems) = lngUserID Then
        matchUserID = intItems
        Exit Function
    End If
  Next intItems
  matchUserID = -1
End Function

Public Function disableDisplay()
  frmProjectMaintenance.txtProjectNumber.Enabled = False
  frmProjectMaintenance.txtProjectName.Enabled = False
  frmProjectMaintenance.txtProjectDescription.Enabled = False
  frmProjectMaintenance.txtDateCreated.Enabled = False
  frmProjectMaintenance.txtDateClosed.Enabled = False
  frmProjectMaintenance.cboCreatedBy.Enabled = False
  frmProjectMaintenance.cboCreatedBy.Enabled = False
  frmProjectMaintenance.cboManager.Enabled = False
  frmProjectMaintenance.txtEstimatedLabour.Enabled = False
  frmProjectMaintenance.txtEstimatedMaterial.Enabled = False
  frmProjectMaintenance.txtEstimatedTravel.Enabled = False
End Function

Public Function enableDisplay()
  frmProjectMaintenance.txtProjectNumber.Enabled = True
  frmProjectMaintenance.txtProjectName.Enabled = True
  frmProjectMaintenance.txtProjectDescription.Enabled = True
  frmProjectMaintenance.txtDateCreated.Enabled = True
  frmProjectMaintenance.txtDateClosed.Enabled = True
  frmProjectMaintenance.cboCreatedBy.Enabled = True
  frmProjectMaintenance.cboCreatedBy.Enabled = True
  frmProjectMaintenance.cboManager.Enabled = True
  frmProjectMaintenance.txtEstimatedLabour.Enabled = True
  frmProjectMaintenance.txtEstimatedMaterial.Enabled = True
  frmProjectMaintenance.txtEstimatedTravel.Enabled = True
End Function


