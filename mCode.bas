Attribute VB_Name = "mCode"
Option Explicit
Public VBP As VBProject
Public VBC As VBComponent
Public Const APP_NAME As String = "Procedure Builder"
Public cpActCodePane As vbide.CodePane     ' To store the active code pane
Public cmCodeModule As vbide.CodeModule    ' To store the active code module
