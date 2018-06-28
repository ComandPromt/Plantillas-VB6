<%@ Register TagPrefix="msdn" TagName="debug" Src="MyDebugTool.ascx" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="WebForm1.aspx.vb" Inherits="TestTracer.WebForm1"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>WebForm1</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio.NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body bgcolor="ivory" style="FONT-SIZE:x-small;FONT-FAMILY:verdana">
		<form runat="server" ID="theForm">
			<msdn:debug runat="server" userkey="dino" ID="MyTracer" />
			<table>
				<tr>
					<td><asp:label runat="server" Text="Key:" ID="Label1" /></td>
					<td><asp:textbox runat="server" Font-Bold="true" id="theKey" /></td>
				</tr>
				<tr>
					<td><asp:label runat="server" Text="Value:" ID="Label2" /></td>
					<td><asp:textbox runat="server" id="theValue" /></td>
				</tr>
			</table>
			<asp:linkbutton runat="server" Text="Add to Cache..." ID="Linkbutton1" />
			<br>
			<br>
		</form>
	</body>
</HTML>
