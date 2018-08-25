Attribute VB_Name = "Module1"
Public blnSkipEvent As Integer

Sub Realize_GraphSettings(intCurrentTab As Integer)

'---------------------------------------------------------------------
' Upon switching tabs in the tabbed toolbar, the new tab needs updated
' information on the current status of the graph.  This routine
' performs the update for the tab in question.
'---------------------------------------------------------------------

  blnSkipEvent = True
  Select Case intCurrentTab
    '-----------------------------------------------------------------
    ' Tab 1  Graph Type information
    '-----------------------------------------------------------------
    Case 0
      frmGraphApp.cboGRTS(0).ListIndex = mdiGraph.graSample.GraphType
      frmGraphApp.cboGRTS(1).ListIndex = mdiGraph.graSample.GraphStyle
      frmGraphApp.txtPointData(0).Text = mdiGraph.graSample.NumSets
      frmGraphApp.txtPointData(1).Text = mdiGraph.graSample.NumPoints
      frmGraphApp.optGridP(mdiGraph.graSample.GridStyle) = True
    '-----------------------------------------------------------------
    ' Tab 2  Data information for graph
    '-----------------------------------------------------------------
    Case 1
      frmGraphApp.hsbPositionBar(0).Value = 1
      frmGraphApp.hsbPositionBar(1).Value = 1
      frmGraphApp.hsbPositionBar(2).Value = 1
      frmGraphApp.hsbPositionBar(0).Max = mdiGraph.graSample.NumSets
      frmGraphApp.hsbPositionBar(1).Max = mdiGraph.graSample.NumPoints
      If (CSng(frmGraphApp.txtPointData(0).Text) >= CSng(frmGraphApp.txtPointData(1).Text)) Then
        frmGraphApp.hsbColorBar.Max = CSng(frmGraphApp.txtPointData(0).Text)
        frmGraphApp.hsbColorBar.Value = mdiGraph.graSample.ThisPoint
        frmGraphApp.hsbPositionBar(1).Max = CInt(frmGraphApp.txtPointData(0).Text)
        frmGraphApp.hsbPositionBar(1).Value = mdiGraph.graSample.ThisPoint
        frmGraphApp.hsbPositionBar(2).Max = CSng(frmGraphApp.txtPointData(0).Text)
       Else
        frmGraphApp.hsbPositionBar(1).Max = CSng(frmGraphApp.txtPointData(1).Text)
        frmGraphApp.hsbPositionBar(2).Max = CSng(frmGraphApp.txtPointData(1).Text)
      End If
    '-----------------------------------------------------------------
    ' Tab 2  Color information for the graph
    '-----------------------------------------------------------------
    Case 2
      frmGraphApp.optColor_Item(1).Value = True
      frmGraphApp.hsbColorBar.Value = 1
      frmGraphApp.optPalt(mdiGraph.graSample.Palette).Value = True
    '-----------------------------------------------------------------
    ' Tab 3  Graph titles and captions
    '-----------------------------------------------------------------
    Case 3
      frmGraphApp.txtTitles(0).Text = mdiGraph.graSample.BottomTitle
      frmGraphApp.txtTitles(1).Text = mdiGraph.graSample.GraphTitle
      frmGraphApp.txtTitles(2).Text = mdiGraph.graSample.LeftTitle
      frmGraphApp.txtTitles(3).Text = mdiGraph.graSample.GraphCaption
      frmGraphApp.cboFontInfo(0).ListIndex = 4
    '-----------------------------------------------------------------
    ' Tab 4  Graph Axis information
    '-----------------------------------------------------------------
    Case 4
      frmGraphApp.cboAx_Props(0).ListIndex = mdiGraph.graSample.YAxisPos
      frmGraphApp.cboAx_Props(1).ListIndex = mdiGraph.graSample.YAxisStyle
      frmGraphApp.cboAx_Props(2).ListIndex = mdiGraph.graSample.Ticks
      frmGraphApp.txtYAxv(0).Text = CStr(mdiGraph.graSample.YAxisMin)
      frmGraphApp.txtYAxv(1).Text = CStr(mdiGraph.graSample.YAxisMax)
      frmGraphApp.txtYAxv(2).Text = CStr(mdiGraph.graSample.YAxisTicks)
      frmGraphApp.txtYAxv(3).Text = CStr(mdiGraph.graSample.TickEvery)
      frmGraphApp.txtYAxv(4).Text = CStr(mdiGraph.graSample.LabelEvery)
    '-----------------------------------------------------------------
    ' Tab 6  Miscellaneous Graph information
    '-----------------------------------------------------------------
    Case 5
      frmGraphApp.cboLineStats.ListIndex = mdiGraph.graSample.LineStats
      frmGraphApp.cboPrintStyle.ListIndex = mdiGraph.graSample.PrintStyle
      frmGraphApp.chkPtln.Value = mdiGraph.graSample.PatternedLines
      frmGraphApp.chkThln.Value = mdiGraph.graSample.ThickLines
  End Select
  blnSkipEvent = False

End Sub

