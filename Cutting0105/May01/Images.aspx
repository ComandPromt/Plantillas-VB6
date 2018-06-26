<%@ Page Language="C#" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<html>
<title>Adapting to Data - Images</title>
<style>
  hr		{height:2px;color:black;}
  .StdText	{font-family:verdana;font-size:9pt;font-weight:bold;}
  .StdTextBox	{font-family:verdana;font-size:9pt;border:solid 1px black;filter:progid:DXImageTransform.Microsoft.dropshadow(OffX=2, OffY=2, Color='gray', Positive='true');}
  .Shadow	{filter:progid:DXImageTransform.Microsoft.dropshadow(OffX=2, OffY=2, Color='gray', Positive='true');}
</style> 


<script runat="server">
public void Page_Load(Object sender, EventArgs e)
{
	// Initialize only the first time...
	if (!Page.IsPostBack)
	{
		lblURL.Text = Request.Url + "<hr>";
	}
}

public void OnLoadData(Object sender, EventArgs e)
{
	SqlConnection conn = new SqlConnection(txtConn.Text);
	SqlDataAdapter da = new SqlDataAdapter(txtCommand.Text, conn);
	
	DataSet ds = new DataSet();
	da.Fill(ds, "MyTable");

	// Bind the data
	grid.DataSource = ds.Tables["MyTable"];

	// Display the data
	grid.DataBind();
}

String GetProperGifFile(int bossID)
{
	if (bossID != 0)
		return "checked.gif";
	return "unchecked.gif";
}

</script>


<body bgcolor="ivory" style="font-family:arial;font-size:9pt">

<!-- ASP.NET topbar -->
<h2>Adapting to Data - Images</h2>
<asp:Label runat="server" cssclass="StdText" font-bold="true">Current path: </asp:label>
<asp:Label runat="server" id="lblURL" cssclass="StdText" style="color:blue"></asp:label>

<form runat="server">

  <table>
  <tr>
  <td><asp:label runat="server"  text="Connection String" cssclass="StdText" /></td>
  <td><asp:textbox runat="server" id="txtConn"
	Enabled="false"
 	cssclass="StdTextBox"
	width="700px"
	text="DATABASE=Northwind;SERVER=localhost;UID=sa;PWD=;" /></td></tr>    

  <tr>
  <td><asp:label runat="server"  text="Command Text" cssclass="StdText"/></td>
  <td><asp:textbox runat="server" id="txtCommand" 
        Enabled="false"
	width="700px"
 	cssclass="StdTextBox"
	text="SELECT employeeid, titleofcourtesy, firstname, lastname, title, ISNULL(reportsto,0) AS boss FROM Employees" /></td></tr></table>    

    <br><br>	
    <asp:linkbutton runat="server" id="btnLoad" text="Go get data..." onclick="OnLoadData" />

    <hr>

    <asp:DataGrid id="grid" runat="server"  
	AutoGenerateColumns="false"
	CssClass="Shadow" BackColor="white"
	CellPadding="2" CellSpacing="0" 
	BorderStyle="solid" BorderColor="black" BorderWidth="1"
	font-size="x-small" font-names="verdana">

	<AlternatingItemStyle BackColor="palegoldenrod" />
	<ItemStyle BackColor="beige" />
	<HeaderStyle ForeColor="white" BackColor="brown" Font-Bold="true" />

        <columns>
	   <asp:BoundColumn runat="server" HeaderText="ID" DataField="employeeid">		
		<itemstyle backcolor="lightblue" font-bold="true" />
	   </asp:BoundColumn>

	   <asp:TemplateColumn runat="server" HeaderText="Employee Name">		
		<itemtemplate>
			<asp:label runat="server" 
				style="margin-left:5;margin-right:5"
				Text='<%# DataBinder.Eval(Container.DataItem, "TitleOfCourtesy") + "<b> " + 
					  DataBinder.Eval(Container.DataItem, "LastName") + "</b>" + ", " + 
					  DataBinder.Eval(Container.DataItem, "FirstName") %>' />			
		</itemtemplate>
	   </asp:TemplateColumn>

	   <asp:TemplateColumn HeaderText="Reports" 
		headerstyle-horizontalalign="Center" 
		itemstyle-horizontalalign="Center">
		
		<itemtemplate>
			<asp:image runat="server" 
				imageurl='<%# GetProperGifFile((int)DataBinder.Eval(Container.DataItem, "boss")) %>'  />
		</itemtemplate>
	   </asp:TemplateColumn>

	   <asp:BoundColumn runat="server" DataField="title" HeaderText="Position" />
 	</columns>
     </asp:DataGrid>

</form>

</body>
</html>
