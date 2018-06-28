<%@ Page Language="C#" Trace="true" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data" %>

<html><head>
<style>
.MyButton	{border:solid groove 1px black;color:black;background:gainsboro;font-family:verdana;font-size:8pt;behavio

</style>

<SCRIPT runat="server"> 
public void Page_Load(Object sender, EventArgs e)
{
	if (!Page.IsPostBack)
	{
		Repeater1.DataSource = (ICollection) CreateDataSource();
		Repeater1.DataBind();
	}
}

public ICollection CreateDataSource()
{
	DataTable dt = new DataTable();
	DataColumn colName = new DataColumn();
	colName.DataType = System.Type.GetType("System.String");
	colName.ColumnName = "FolderName";
	dt.Columns.Add(colName);
	
	DataColumn colDesc = new DataColumn();
	colDesc.DataType = System.Type.GetType("System.String");
	colDesc.ColumnName = "FolderDesc";
	dt.Columns.Add(colDesc);          

	String strDir = "C:\\Documents and Settings\\Administrator\\My Documents\\My Pictures";
	Directory dirMP = new Directory(strDir);
	foreach (Directory d in dirMP.GetDirectories())
	{
		DataRow dr = dt.NewRow();
		dr["FolderName"] = d.Name;
		dr["FolderDesc"] = "Content of " + d.Name;
		dt.Rows.Add(dr);
 	}
 	
 	return dt.DefaultView;
}

public void OnFolderSelected(Object sender, EventArgs e)
{
	String strDir = "C:\\Documents and Settings\\Administrator\\My Documents\\My Pictures";
	Button b = (Button) sender;
	fViewPictures.Attributes["src"] = strDir + "\\" + b.Text;
	Trace.Write("ID", b.ID.ToString());
}


</SCRIPT> 

<BODY bgcolor="ivory" style="font-family:verdana;font-size:9px">
<form runat="server">

<asp:repeater runat="server" id="Repeater1">
	
<template name="HeaderTemplate">
<h2>Pictures Folders</h2>
   <table border="0" cellspacing="8"><tr>
</template>

<template name="ItemTemplate">
	<td valign="top" width="130px">
  		<asp:button runat="server" width="100%" cssclass="MyButton" id="btnFolder" OnClick="OnFolderSelected"
  		  	Text='<%# DataBinder.Eval(Container.DataItem, "FolderName") %>' />
  		<br> 
  		<asp:label runat="server" Font-size="10px"
  			Text='<%# DataBinder.Eval(Container.DataItem, "FolderDesc")%>' />
	</td>
</template>

<template name="FooterTemplate">
   </tr></table>
   <hr>
</template>

</asp:Repeater>

<iframe runat="server" id="fViewPictures" src="c:\" width="100%" height="400px" />
</form>
</body>
</html>