<%@ Page Language="C#" %>
<%@ Register TagPrefix="expo" TagName="ViewPanel" Src="EmployeeViewForm.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<html>
<title>Employees Manager</title>
<style>
  a		{behavior:url(..\..\mouseover.htc);}
  hr		{height:2px;color:black;}
  .StdText	{font-family:verdana;font-size:x-small;}
  .StdTextBox	{font-family:verdana;font-size:x-small;border:solid 1px black;filter:progid:DXImageTransform.Microsoft.dropshadow(OffX=2, OffY=2, Color='gray', Positive='true');}
  .FlatButton	{font-family:verdana;font-size:x-small;border:solid 1px black;behavior:url(..\..\mouseover.htc?ForeColor="blue");}
  .Shadow	{filter:progid:DXImageTransform.Microsoft.dropshadow(OffX=2, OffY=2, Color='gray', Positive='true');}
</style> 


<script runat="server">
public void Page_Load(Object sender, EventArgs e)
{
	// Initialize only the first time...
	if (!Page.IsPostBack)
	{
		lblURL.Text = Request.Url + "<hr>";
		LoadData();
		UpdateView();
	}
}

public void PageIndexChanged(Object sender, DataGridPageChangedEventArgs e) 
{
	// Set the current item to edit mode
	grid.CurrentPageIndex = e.NewPageIndex;

	// Refresh the grid
	UpdateView();
}

public void ItemCreated(Object sender, DataGridItemEventArgs e)
{
	ListItemType lit = e.Item.ItemType;

	if (lit == ListItemType.Pager) 
	{
		// The pager as a whole has the following layout:
		//
		// <TR><TD colspan=X> ... links ... </TD></TR> 
		//
		// Item points to <TR>. The code below moves to <TD>.
		TableCell pager = (TableCell) e.Item.Controls[0];

		// Loop through the pager buttons skipping over blanks
		// (Blanks are treated as LiteralControl(s)
		for (int i=0; i<pager.Controls.Count; i+=2) 
		{
			Object o = pager.Controls[i];
			if (o is LinkButton) 
			{
				LinkButton h = (LinkButton) o;
				h.Text = "[ " + h.Text + " ]"; 
			}
			else
			{
				Label l = (Label) o;
				l.Text = "<b>Page " + l.Text + "</b>"; 
			}
		}
	}
}


public void SelectionIndexChanged(Object sender, EventArgs e) 
{
	ToggleBitmap(grid.SelectedItem);
	SelectRecord();
}


public void OnSelectRecordByID(Object sender, EventArgs e)
{
	int nEmpID = Convert.ToInt32(txtEmployeeID.Text);
	grid.SelectedIndex = GetPageIndexFromID(nEmpID);
	btnUnselect.Enabled = true;	

	SelectRecordByID(nEmpID);
	ToggleBitmap(grid.SelectedItem);
}

public void OnApplyFilter(Object sender, EventArgs e)
{
	DataSet ds = (DataSet) Session["MyData"];
	try {
		DataViewManager dvm = ds.DefaultViewManager;
		dvm.DataViewSettings["MyTable"].RowFilter = txtFilterString.Text;
 		grid.DataSource = ds;
		grid.DataBind();
		Statusbar.Text = "";
	} catch(Exception exc) {
		Statusbar.Text = "<b>There were errors applying the filter: </b>" + exc.Message;
	}

}

public void OnUnselectRecord(Object sender, EventArgs e)
{
	UnselectRecord();
}


////////////////////////////////////////////////////////////////////////

private void LoadData()
{
	String strCnn = "DATABASE=Northwind;SERVER=localhost;Integrated Security=SSPI;";
	String strCmd = "SELECT * FROM Employees";

	SqlConnection conn = new SqlConnection(strCnn);
	SqlDataAdapter da = new SqlDataAdapter(strCmd, conn);
	DataSet ds = new DataSet();
	da.Fill(ds, "MyTable");

	Session["MyData"] = ds;
}

private void UpdateView()
{
	DataSet ds = (DataSet) Session["MyData"];

	// Bind the data
	grid.DataSource = ds.Tables["MyTable"];

	// Display the data
	grid.DataBind();
}

private void SelectRecord()
{
	btnUnselect.Enabled = true;
	SelectRecordByID((int) grid.DataKeys[grid.SelectedIndex]);
}

private void UnselectRecord()
{
	btnUnselect.Enabled = false;		
	view.ClearAll(); 
	grid.SelectedIndex = -1;
}

private void SelectRecordByID(int nEmpID)
{
	DataSet ds = (DataSet) Session["MyData"];
	DataTable dt = ds.Tables["MyTable"];

	// There seems to be no way to make a search within the filtered view. Work around this by
	// ANDing the filter string (if any) with the SELECT expression.
	String strFilter = "";
	if (bSearchOnFilter.Checked)
		strFilter = (txtFilterString.Text != "" ?txtFilterString.Text + " AND " :"");

	DataRow[] a = dt.Select(strFilter + "EmployeeID=" + nEmpID.ToString());
	try {
	   view.EmployeeID = a[0]["EmployeeID"].ToString();
	   view.TitleOfCourtesy = a[0]["TitleOfCourtesy"].ToString();
	   view.FirstName = a[0]["firstname"].ToString();
	   view.LastName = a[0]["lastname"].ToString();
	   view.Title = a[0]["title"].ToString();
	   view.HireDate = ((DateTime)a[0]["hiredate"]).ToShortDateString();
	   view.Notes = a[0]["notes"].ToString();
           Statusbar.Text = "";
	}
	catch (Exception exc) {
                   view.ClearAll(); 
                   Statusbar.Text = "<b>There were errors searching for the record : </b>" + exc.Message;
	}
}

private int GetPageIndexFromID(int nEmpID)
{
	int nRetValue = -1;

	for (int i=0; i<grid.DataKeys.Count; i++)
		if (nEmpID == (int) grid.DataKeys[i])
		{
			nRetValue = i;
			break;
		}

	return nRetValue;	
}

private void ToggleBitmap(DataGridItem dgi)
{
	// Change the bitmap on the "first" column to a pinned button
	TableCell c = (TableCell) dgi.Controls[0];
	c.Text = "<img src=selected.gif border=0 align=absmiddle>";	
}
</script>


<body bgcolor="ivory" style="font-family:arial;font-size:x-small">

<!-- ASP.NET topbar -->
<h2>Employees Manager</h2>
<asp:Label runat="server" cssclass="StdText" font-bold="true">Current path: </asp:label>
<asp:Label runat="server" id="lblURL" cssclass="StdText" style="color:blue"></asp:label>

<form runat="server">
<table><tr>
<td valign="top">
    <asp:DataGrid id="grid" runat="server"  
	AutoGenerateColumns="false"
	CssClass="Shadow" BackColor="white"
	CellPadding="2" CellSpacing="0" 
	BorderStyle="solid" BorderColor="black" BorderWidth="1"
	Font-Size="x-small" Font-Names="verdana"
	AllowPaging="true"
	PageSize="4"
	DataKeyField="employeeid"
	OnSelectedIndexChanged="SelectionIndexChanged"
	OnItemCreated="ItemCreated"
	OnPageIndexChanged="PageIndexChanged">

	<AlternatingItemStyle BackColor="palegoldenrod" />
	<ItemStyle BackColor="beige" />
	<PagerStyle Mode="NumericPages" HorizontalAlign="right" />
	<SelectedItemStyle ForeColor="white" BackColor="#77B4EE" Font-Bold="true" />
	<HeaderStyle ForeColor="white" BackColor="brown" HorizontalAlign="center" Font-Bold="true" />

        <columns>
	   <asp:ButtonColumn CommandName="select" 
			Text="<img border=0 alt='Select' align=absmiddle src=unselected.gif>" />

	   <asp:TemplateColumn runat="server" HeaderText="Employee Name">		
		<itemtemplate>
			<asp:label runat="server" 
				style="margin-left:5;margin-right:5"
				Text='<%# DataBinder.Eval(Container.DataItem, "TitleOfCourtesy") + "<b> " + 
					  DataBinder.Eval(Container.DataItem, "LastName") + "</b>" + ", " + 
					  DataBinder.Eval(Container.DataItem, "FirstName") %>' />			
		</itemtemplate>
	   </asp:TemplateColumn>

	   <asp:BoundColumn runat="server" DataField="title" HeaderText="Position" />
	   <asp:BoundColumn runat="server" DataField="country" HeaderText="From" />
 	</columns>
     </asp:DataGrid>
</td>
<td valign="top">
	<expo:ViewPanel runat="server" id="view" />
	<hr>
	<asp:label runat="server" text="<b>ID</b>" />
	<asp:textbox runat="server" id="txtEmployeeID" text="1" width="50px" /> 
	<asp:linkbutton runat="server" 
		text="Go" 
		tooltip="Select matching record"
		onclick="OnSelectRecordByID" />
</td>
</tr></table>

<asp:linkbutton runat="server" id="btnUnselect" enabled="false" 
	text="Unselect" cssclass="stdtext" onclick="OnUnselectRecord" />
<hr>

<asp:label runat="server" cssclass="StdText" text="Filter: " Font-Bold="true" />
<asp:textbox runat="server" cssclass="StdTextBox" id="txtFilterString" width="400px"  text="EmployeeID>3" />
<asp:button runat="server" cssclass="FlatButton" text="Apply" onclick="OnApplyFilter" />
&nbsp;&nbsp;&nbsp;&nbsp;
<asp:checkbox runat="server" id="bSearchOnFilter" text="Restrict search to the current filtered view" />
<hr>
<asp:label runat="server" ForeColor="red" cssclass="StdText" id="Statusbar" />

</form>

</body>
</html>
