<%@ Page Language="C#" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data" %>

<html>
<title>Listing directories</title>
<style>
  hr		{height:2px;color:black;}
  .StdText	{font-family:verdana;font-size:9pt;}
  .StdTextBox	{font-family:verdana;font-size:9pt;border:solid 1px black;}
</style> 

<SCRIPT runat="server"> 
public void Page_Load(Object sender, EventArgs e)
{
	if (!Page.IsPostBack)
		lblURL.Text = Request.Url + "<hr>";
}

public ICollection CreateDataSource(String strDir)
{
	DataTable dt = new DataTable();
	DataColumn colName = new DataColumn();
	colName.DataType = System.Type.GetType("System.String");
	colName.ColumnName = "FolderName";
	dt.Columns.Add(colName);
	
	DataColumn colTime = new DataColumn();
	colTime.DataType = System.Type.GetType("System.String");
	colTime.ColumnName = "FolderTime";
	dt.Columns.Add(colTime);          

	DirectoryInfo dir = new DirectoryInfo(strDir);
	foreach (DirectoryInfo d in dir.GetDirectories())
	{
		DataRow dr = dt.NewRow();
		dr["FolderName"] = d.Name;
		dr["FolderTime"] = d.CreationTime.ToString();
		dt.Rows.Add(dr);
 	}
 	
 	return dt.DefaultView;
}

public void RetrieveDirectories(Object sender, EventArgs e)
{
	try {
		statusbar.Text = "";
		Repeater1.DataSource = (ICollection) CreateDataSource(theDir.Text);
	} catch {
		statusbar.Text = "Directory not found.";
	}
	Repeater1.DataBind();
}

</SCRIPT> 

<BODY bgcolor="ivory" style="font-family:verdana;font-size:9pt">
<form runat="server">

<!-- ASP.NET topbar -->
<h2>Listing directories</h2>
<asp:Label runat="server" cssclass="StdText" font-bold="true">Current path: </asp:label>
<asp:Label runat="server" id="lblURL" cssclass="StdText" style="color:blue"></asp:label>

<!-- Enter the parent directory -->
Enter the directory name:
<asp:textbox runat="server" id="theDir" cssclass="StdTextBox" />
<asp:linkbutton runat="server" id="btnExecute" text="Go" onclick="RetrieveDirectories" />
<hr>

<asp:repeater runat="server" id="Repeater1">
	
<HeaderTemplate>
   <table style="border:1px solid black" class="stdtext">
      <thead bgcolor="blue" style="color:white">
         <td><b>Folder Name</b></td>
         <td><b>Creation Time</b></td>
      </thead>
</HeaderTemplate>

<ItemTemplate>
   <tr>
      <td bgcolor="white"> <%# ((DataRowView)Container.DataItem)["FolderName"] %> </td>
      <td bgcolor="white"> <%# DataBinder.Eval(Container.DataItem, "FolderTime") %> </td>
   </tr>
</ItemTemplate>

<AlternatingItemTemplate>
   <tr>
      <td bgcolor="lightblue"> <%# DataBinder.Eval(Container.DataItem, "FolderName") %> </td>
      <td bgcolor="lightblue"> <%# DataBinder.Eval(Container.DataItem, "FolderTime") %> </td>
   </tr>
</AlternatingItemTemplate>

<FooterTemplate>
   <tfoot>
	<td bgcolor="silver" colspan=2><%# "<b>" + ((DataView)Repeater1.DataSource).Count + "</b> directories found."%></td>
   </tfoot>	
   </table>
</FooterTemplate>

</asp:Repeater>

<!-- Messages go here -->
<asp:label runat="server" Font-Italic="true" id="statusbar" />
</form>
</body>
</html>