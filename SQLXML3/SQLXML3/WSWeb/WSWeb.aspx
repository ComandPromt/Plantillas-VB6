<%@ Page language="c#" Codebehind="WSWeb.aspx.cs" AutoEventWireup="false" Inherits="WSWeb.WSWebForm" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>Web Services Web Client</title>
		<meta name="GENERATOR" Content="Microsoft Visual Studio 7.0">
		<meta name="CODE_LANGUAGE" Content="C#">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout" bgColor="#99ccff">
		<form id="frmWSWeb" method="post" runat="server">
			<asp:Button id="cmdExecute" style="Z-INDEX: 101; LEFT: 820px; POSITION: absolute; TOP: 47px" runat="server" Text="Execute" Width="178px"></asp:Button>
			<asp:DataGrid id="DataGrid2" style="Z-INDEX: 110; LEFT: 25px; POSITION: absolute; TOP: 423px" runat="server" Width="972px" Height="111px" Font-Size="Small" BackColor="White" BorderColor="#400000" AllowSorting="True" AllowPaging="True" BorderWidth="2px">
				<HeaderStyle BorderStyle="Outset"></HeaderStyle>
				<PagerStyle BorderStyle="Ridge"></PagerStyle>
			</asp:DataGrid>
			<asp:Label id="Label1" style="Z-INDEX: 102; LEFT: 35px; POSITION: absolute; TOP: 58px" runat="server">Customer ID:</asp:Label>
			<asp:TextBox id="txtParm1" style="Z-INDEX: 105; LEFT: 129px; POSITION: absolute; TOP: 56px" runat="server">ALFKI</asp:TextBox>
			<asp:DataGrid id="DataGrid1" style="Z-INDEX: 108; LEFT: 25px; POSITION: absolute; TOP: 116px" runat="server" Width="972px" Height="111px" Font-Size="Small" BackColor="White" BorderColor="#400000" AllowSorting="True" AllowPaging="True" BorderWidth="2px">
				<HeaderStyle BorderStyle="Outset"></HeaderStyle>
				<PagerStyle BorderStyle="Ridge"></PagerStyle>
			</asp:DataGrid>
			<asp:Label id="Label4" style="Z-INDEX: 109; LEFT: 34px; POSITION: absolute; TOP: 13px" runat="server" Width="459px" Height="14px" Font-Bold="True" Font-Size="Large">Web Services Web-Based Test Harness</asp:Label>
			<asp:Label id="Label5" style="Z-INDEX: 111; LEFT: 34px; POSITION: absolute; TOP: 397px" runat="server">All Customers</asp:Label>
			<asp:Label id="Label6" style="Z-INDEX: 112; LEFT: 35px; POSITION: absolute; TOP: 90px" runat="server">Customer Order History</asp:Label>
		</form>
	</body>
</HTML>
