<%@ Page Language="C#" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SQL" %>
<html><head>

<style>
.PagerSpan {font-weight:bold;color:blue;}
.PagerLink {behavior:url(MouseOver.htc);}
</style>

<script runat="server">
public void Page_Load(Object Sender, EventArgs e)
{
	SearchData(Sender, e);
}

/*
	This function creates the data source whenever is 
	necessary. If you're going to use paging you do always 
	need to reload the data source before re-binding. 
	You cannot cache the dataset in this case.
*/
private ICollection CreateDataSource()
{
	// Set up the SQL2K connection
	String strConn;
	strConn = "DATABASE=Northwind;SERVER=localhost;UID=sa;PWD=;";

	// Set the SQL command to run
	String strCmd = "";
	strCmd += "SELECT ";
	strCmd += " employeeid, ";
	strCmd += "titleofcourtesy + ' ' + firstname + ' ' + lastname AS EmployeeName, ";
	strCmd += "title ";
	strCmd += "FROM Employees";
	
	// Execute the command and add a named table to the dataset
	SQLDataSetCommand oCMD = new SQLDataSetCommand(strCmd, strConn);

	// Add a named table to the dataset 
	DataSet oDS = new DataSet();
	oCMD.FillDataSet(oDS, "EmployeesList");
	
	// Return a given data table
	return oDS.Tables["EmployeesList"].DefaultView;
}


/*
	The paging needs this delegate handler to refresh the
	data source. It's critical that you recalculate the data
	source here. Internally, the datagrid will load only the 
	needed elements.
*/
public void Grid_Change(Object Sender, DataGridPageChangedEventArgs e)
{
	DataGrid1.DataSource = CreateDataSource();
	DataGrid1.DataBind();
}

/*
	This function gets called to fill the grid upon page loading.
*/
public void SearchData(Object Sender, EventArgs e)
{
	SetGridProperties();
	DataGrid1.DataSource = CreateDataSource();
	DataGrid1.DataBind();
}

/*
	This function makes sure the datagrid reflects some properties 
	set dynamically.
*/
private void SetGridProperties()
{
	DataGrid1.PagerStyle.Mode = PagerMode.NumericPages;
	DataGrid1.PagerStyle.BackColor = System.Drawing.Color.Gainsboro;
	DataGrid1.PagerStyle.PageButtonCount = 10;
	DataGrid1.PageSize = 4;	
}


/*
	This function gets called whenever a new item is created within
	the datagrid. You check the type, enumerate the pager items
	and set the CSS style
*/
public void Grid_NewItem(Object Sender, DataGridItemEventArgs e)
{
	// Get the newly created item
	ListItemType itemType = e.Item.ItemType;
	
	// Is it the pager?
    if (itemType == ListItemType.Pager) 
    {
		// There's just one control in the list...
		TableCell pager = (TableCell)e.Item.Controls[0];
		
		// Enumerates all the items in the pager...
		for (int i=0; i<pager.Controls.Count; i+=2) 
		{
			try {
			  Label l = (Label) pager.Controls[i];
			  l.Text = "Page " + l.Text;
			  l.CssClass = "PagerSpan";
			} 
			catch {
			  LinkButton h = (LinkButton) pager.Controls[i];
			  h.Text = "[ " + h.Text + " ]";
			  h.CssClass = "PagerLink";
			}
		}
    }
}

/*
	This code runs whenever an item is selected within the grid.
*/
public void Grid_SelectionChanged(Object Sender, EventArgs e) 
{
	try {
	DataGridItem dgi = DataGrid1.SelectedItem;
	TableCell cName = (TableCell) dgi.Controls[2];
	TableCell cTitle = (TableCell) dgi.Controls[3];
	statusbar.Text = "<b>You have selected </b>" + cName.Text + " (" + cTitle.Text + ")" + "<br><b> Item:</b> " + DataGrid1.SelectedIndex.ToString();
	TableCell c = (TableCell) dgi.Controls[0];
	c.Text = "<img src=opened.gif border=0 align=absmiddle>";	
	}
	catch {
	statusbar.Text = "<b>You have no item selected </b><br><b> Item:</b> " + DataGrid1.SelectedIndex.ToString();
	}
}

void SelectRow(Object sender, EventArgs e) 
{
	DataGrid1.SelectedIndex = RowNumber.Text.ToInt32()-1;
	Grid_SelectionChanged(sender, e);
}


</script>
</head>
<body>

<%
	Session["ShowSampleBar"] = "no";
	Session["Logo"] = "../../images/expoware.gif";
	/*
	Session["FileName"] = Request.ServerVariables["SCRIPT_NAME"];
	Session["Day"] = 2;
	Session["Course"] = "Programming ADO.NET";
	Session["Sample"] = "Step 5";
	Session["SampleDesc"] = "Do something with ASP.NET data-binding";
	*/
	Server.Execute("../../layout.aspx");
%>

<!-- ASP.NET Pages can contain just one server-side form -->
<form id=Form1 runat="server">
<hr noshade Style="HEIGHT: 10px" >

<!-- Show the information -->&nbsp; 
<asp:datagrid id=DataGrid1 runat="server" 
	AutoGenerateColumns = False 
	CellPadding="2" 
	CellSpacing="2" 
	GridLines="None" 
	BorderStyle="solid"
	BorderColor="black"
	BorderWidth="1"
	ForeColor="Black" 
	font-size="x-small" 
	font-names="Verdana" 
	width="100%" 
	PagerStyle-HorizontalAlign="Right"
	AllowPaging="True" 
	OnItemCreated="Grid_NewItem"
	OnPageIndexChanged="Grid_Change"
	OnSelectedIndexChanged="Grid_SelectionChanged">
<property name="Columns">
<asp:ButtonColumn 
CommandName="Select" 
Text="<img border=0 align=absmiddle src=closed.gif>"></asp:ButtonColumn>
<asp:BoundColumn 
headerstyle-horizontalalign="Center" DataField="employeeid" 
HeaderText="ID"></asp:BoundColumn>
<asp:BoundColumn 
headerstyle-horizontalalign="Center" DataField="EmployeeName" 
HeaderText="Name"></asp:BoundColumn>
<asp:BoundColumn 
headerstyle-horizontalalign="Center" DataField="title" 
HeaderText="Occupation"></asp:BoundColumn></property><property 
name="AlternatingItemStyle">
<asp:TableItemStyle 
BackColor="palegreen"></asp:TableItemStyle></property><property 
name="ItemStyle">
<asp:TableItemStyle 
BackColor="beige"></asp:TableItemStyle></property><property 
name="HeaderStyle">
<asp:TableItemStyle ForeColor="White" BackColor="brown" 
Font-Bold="True"></asp:TableItemStyle></property><property 
name="SelectedItemStyle">
<asp:TableItemStyle ForeColor="white" BackColor="blue" 
Font-Bold="true"></asp:TableItemStyle></property>

</asp:datagrid>

<br>
<asp:label width="100%" height="30" backcolor="#ffcc00"  font-size="x-small" font-names="Verdana" runat=server id="statusbar"></asp:label>
<asp:Label id=Label1 runat="server">Enter the row to 
select</asp:Label> 

	<asp:TextBox id="RowNumber" runat="server"
		CssClass= "PagerText"
		maxlength ="2"
		Text="1" />
        
	<asp:button id="SelectButton" runat="server"    
		Text="Select"
		CssClass = "PagerPush"
        CommandArgument = "goto"
        OnClick = "SelectRow" />

</form>
</body></html>
