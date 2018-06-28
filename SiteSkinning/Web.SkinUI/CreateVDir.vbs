WScript.Echo "IIS VRoot Creation Tool"


Set objArgs = WScript.Arguments

if objArgs.Count <> 1 then 
	EchoHelp
	WScript.Quit
end if

CreateVDir objArgs(0)

sub EchoHelp

  WScript.Echo vbCRLF & "Used to expose a subdirectory as an IIS Virtual Directory"
  WScript.Echo "Usage:" & vbCRLF & "  CreateVDir <<subdirname>>"
  WScript.Echo "Example:" & vbCRLF & "  CreateVDir MySubDirectory"

End Sub

sub CreateVDir(dir)
    Dim oRootNode, RootNodePath, oVirtualDir, oWebSite 

    ' Set the default path 
    RootNodePath = "IIS://LocalHost/w3svc/1" 
    Set oWebSite = GetObject(RootNodePath) 
    If Err <> 0 Then 
        Display "Couldn't get the first node!" 
        WScript.Quit (1) 
    End If 

    Set oRootNode = oWebSite.GetObject("IIsWebVirtualDir", "Root") 
    If Err <> 0 Then 
        Display "Couldn't get the first node!" 
        WScript.Quit (1) 
    End If 


    'Create Virtual Directory 
    Set oVirtualDir = oRootNode.Create("IIsWebVirtualDir", dir) 

    Dim WshShell
    Set WshShell = WScript.CreateObject("WScript.Shell")

    'Set Properties 
    With oVirtualDir 
        .Path = WshShell.CurrentDirectory & "\" & dir
	.AppCreate false
        .SetInfo           		'Save settings to Metabase 
    End With 
end sub