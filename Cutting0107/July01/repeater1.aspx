<%@ Page Language="C#" Trace="false" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data" %>

<html><head>

<SCRIPT runat="server"> 
public void Page_Load(Object sender, EventArgs e)
{
	Repeater1.DataSource = (ICollection) CreateDataSource();
	Repeater1.DataBind();
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

</SCRIPT> 

<BODY bgcolor="ivory" style="font-family:verdana;font-size:9px">
<form runat="server">

<asp:repeater runat="server" id="Repeater1">
	
<template name="HeaderTemplate">
   <table border=1>
      <tr>
         <td><b>Pictures Folder</b></td>
         <td><b>Description</b></td>
      </tr>
</template>

<template name="ItemTemplate">
   <tr>
      <td> <%# DataBinder.Eval(Container.DataItem, "FolderName") %> </td>
      <td> <%# DataBinder.Eval(Container.DataItem, "FolderDesc") %> </td>
   </tr>
</template>

<template name="AlternatingItemTemplate">
   <tr>
      <td bgcolor="lightblue"> <%# DataBinder.Eval(Container.DataItem, "FolderName") %> </td>
      <td bgcolor="lightblue"> <%# DataBinder.Eval(Container.DataItem, "FolderDesc") %> </td>
   </tr>
</template>

<template name="FooterTemplate">
   </table>
</template>

</asp:Repeater>

</form>
</body>
</html>