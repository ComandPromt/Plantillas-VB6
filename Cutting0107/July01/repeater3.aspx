<%@ Page Language="C#" Trace="false" %>
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
	
	lblFolder.Text = b.Text;
	DataList1.DataSource = (ICollection) CreateImageDataSource(strDir + "\\" + b.Text);
	DataList1.DataBind();
}

public ICollection CreateImageDataSource(String strPath)
{
	DataTable dt = new DataTable();
	DataColumn colName = new DataColumn();
	colName.DataType = System.Type.GetType("System.String");
	colName.ColumnName = "ImageName";
	dt.Columns.Add(colName);
	
	DataColumn colDesc = new DataColumn();
	colDesc.DataType = System.Type.GetType("System.String");
	colDesc.ColumnName = "ImageDesc";
	dt.Columns.Add(colDesc);          

	Directory dirMP = new Directory(strPath);
	foreach (File f in dirMP.GetFiles("*.jpg"))
	{
		DataRow dr = dt.NewRow();
		dr["ImageName"] = f.FullName;
		dr["ImageDesc"] = f.Name;
		dt.Rows.Add(dr);
 	}
 	
 	return dt.DefaultView;
}

</SCRIPT> 

<BODY bgcolor="ivory" style="font-family:verdana;font-size:9px">
<form runat="server">

<!-- List the folders available -->
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

<!-- List the content of the selected folder -->
<h2><asp:Label runat="server" id="lblFolder" /></h2>
<asp:DataList id="DataList1" runat="server"
	repeatlayout="table" repeatcolumns="10" repeatdirection="Horizontal">
	
	<property name="SelectedItemStyle">
		<asp:TableItemStyle BorderColor="black" BorderStyle="inset" BorderWidth=3 BackColor="#9BDBFF" />
	</property>
		
	<template name="ItemTemplate">
		<asp:linkbutton font-size="x-small" runat="server" commandname="select" Text="Select" /><br>
		<img align="top" width="90" height="90" border="1" src='<%# DataBinder.Eval(Container.DataItem, "ImageName") %>'
	</template>
	
	<template name="FooterTemplate">
		<hr>
	</template>
</asp:DataList>

</form>
</body>
</html>