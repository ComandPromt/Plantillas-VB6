Attribute VB_Name = "basMain"
'Demonstrates how to use the class
Public Sub Main()

Dim o As CNetworkEnum
Set o = New CNetworkEnum

Call o.SetResourceType(0) 'All
'Call o.SetResourceType(1) 'Machines and shares
'Call o.SetResourceType(2) 'Printers
Call o.Reset

Debug.Print "Directory List"
Debug.Print o.GetDirectoryList

Debug.Print "Domain List"
Debug.Print o.GetDomainList

Debug.Print "File List"
Debug.Print o.GetFileList

Debug.Print "Generic List"
Debug.Print o.GetGenericList

Debug.Print "Group List"
Debug.Print o.GetGroupList

Debug.Print "Local Machine Name"
Debug.Print o.GetLocalMachineName

Debug.Print "Network List"
Debug.Print o.GetNetworkList

Debug.Print "Printer List"
Debug.Print o.GetPrinterList

Debug.Print "Root List"
Debug.Print o.GetRootList

Debug.Print "Server List"
Debug.Print o.GetServerList

Debug.Print "ShareAdmin List"
Debug.Print o.GetShareAdminList

Debug.Print "Share List"
Debug.Print o.GetShareList

Debug.Print "Local User Name"
Debug.Print o.GetLocalUserName

End Sub
