Attribute VB_Name = "userMaintenanceCode"
Public Function disableDisplay()
  frmUserMaintenance.txtFirstName.Enabled = False
  frmUserMaintenance.txtLastName.Enabled = False
  frmUserMaintenance.txtAbbreviation.Enabled = False
  frmUserMaintenance.cboSecurityLevel.Enabled = False
End Function

Public Function enableDisplay()
  frmUserMaintenance.txtFirstName.Enabled = True
  frmUserMaintenance.txtLastName.Enabled = True
  frmUserMaintenance.txtAbbreviation.Enabled = True
  frmUserMaintenance.cboSecurityLevel.Enabled = True
End Function



