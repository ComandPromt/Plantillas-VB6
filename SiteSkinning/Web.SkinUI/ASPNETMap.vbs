WScript.Echo "ASP.NET Mapping Tool"

Set objArgs = WScript.Arguments

if objArgs.Count <> 1 then 
	EchoHelp
	WScript.Quit
end if

MapDefault objArgs(0)

sub EchoHelp

  WScript.Echo vbCRLF & "Usage:" & vbCRLF & _
    "  ASPNETMap <<extension>>"& vbCRLF & vbCRLF & _
    "Example:" & vbCRLF & _
    "  ASPNETMap xml"

End Sub

sub UpdatePath(ext, dllPath, arMaps, i)
	WScript.Echo "  Before: " & arMaps(i)
	arMaps(i) = "." & ext & "," & dllPath & ",1,GET,HEAD,POST,DEBUG"
	WScript.Echo "  After:  " & arMaps(i)
end sub

sub AddPath(ext, dllPath, arMaps)
	if lbound(arMaps) <> 0 then 
		WScript.Echo "LBound = " & lbound(arMaps) & ". Add Extension not supported unless lbound = 0"
		exit sub
	end if
	
	newUpper = ubound(arMaps) + 1
	redim preserve arMaps(newUpper)
	UpdatePath ext, dllPath, arMaps, newUpper
end sub

sub MapWorkhorse(ext, dllPath)
  
  Set oWebRoot = GetObject("IIS://localhost/w3svc/1/root")
  arMaps = oWebRoot.GetEx("ScriptMaps")

  found = -1  
  for i = lbound(arMaps) to UBound(arMaps)
	if LCase(Mid(arMaps(i), 2, len(ext))) = lcase(ext) then 
		found = i
		exit for
	end if
  next
  
  if found <> -1 then UpdatePath ext, dllPath, arMaps, found else AddPath ext, dllPath, arMaps
  
  oWebRoot.PutEx 2, "ScriptMaps", arMaps
  oWebRoot.SetInfo
  
end sub

sub MapDefault(ext)
  On Error Resume Next

  'Get the RootVer of ASP.NET
  Set WshShell = WScript.CreateObject("WScript.Shell")
  rootVer = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\ASP.NET\RootVer")

  'Get the DllFullPath for the RootVer. If successful, pass it to MapWorkhorse
  dllPath = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\ASP.NET\" & rootVer & "\DllFullPath")
  if Err.Number = 0 then 
	MapWorkhorse ext, dllPath
	exit sub
  end if

  'There seems to be a bug with .NET SP2. The registry path uses the new version number
  '(1.0.3705.288) but rootVer is still set to the old version number (1.0.3705.0). 
  'This is a hard coded check for .NET SP2
  Err.Clear
  dllPath = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\ASP.NET\1.0.3705.288\DllFullPath")
  if Err.Number = 0 then 
	MapWorkhorse ext, dllPath
	exit sub
  end if

  WScript.Echo "Could not determine the location of the ASP.NET ISAPI DLL"
end sub
