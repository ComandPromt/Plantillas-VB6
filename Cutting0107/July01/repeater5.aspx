<%@ Page Language="C#" Trace="false" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SQL" %>

<html><head>
<style>
.MyButton	{border:solid groove 1px black;color:black;background:gainsboro;font-family:verdana;font-size:8pt;behavio

</style>

<SCRIPT runat="server"> 
public void Page_Load(Object sender, EventArgs e)
{
	if (!Page.IsPostBack)
	{
		CreateDataSource();
		DataList1.DataSource = (ICollection) GetView();
		DataList1.DataBind();
	}
}

public void CreateDataSource()
{
	String strCmd = "SELECT employeeid, firstname, lastname FROM Employees";
	String strCon = "database=northwind;uid=sa;pwd=;server=localhost;";
	SQLDataSetCommand oCMD = new SQLDataSetCommand(strCmd, strCon);
	DataSet data = new DataSet();
	oCMD.FillDataSet(data, "EmployeesList");
	
	Session["MyTable"] = data.Tables["EmployeesList"];
	Session["RecCount"] = data.Tables["EmployeesList"].Rows.Count;
	Session["CurrentIndex"] = 1;
 }

public ICollection GetView()
{
	DataTable data = (DataTable) Session["MyTable"];
	int nCurrentIndex = (int) Session["CurrentIndex"];
	
	DataView dv = new DataView();
	dv.Table = data;
	
	dv.RowFilter = "EmployeeID=" + data.Rows[nCurrentIndex-1]["employeeid"];
	return dv;
}

public void PrevRecord(Object sender, EventArgs e)
{
	int nCurrentIndex = (int) Session["CurrentIndex"];
	if (nCurrentIndex >1)
		nCurrentIndex --;
	Session["CurrentIndex"] = nCurrentIndex;

	DataList1.DataSource = (ICollection) GetView();
	DataList1.DataBind();
}

public void NextRecord(Object sender, EventArgs e)
{
	int nCurrentIndex = (int) Session["CurrentIndex"];
	int nRecCount = (int) Session["RecCount"];

	if (nCurrentIndex <nRecCount)
	nCurrentIndex ++;
	Session["CurrentIndex"] = nCurrentIndex;

	DataList1.DataSource = (ICollection) GetView();
	DataList1.DataBind();
}

</SCRIPT> 

<BODY bgcolor="ivory" style="font-family:verdana;font-size:9px">
<form runat="server">

<!-- List the folders available -->
<asp:datalist runat="server" id="DataList1">
<template name="HeaderTemplate">
<h2>Employees Folder</h2>
   <table border="0" cellspacing="8"><tr>
</template>

<template name="ItemTemplate">
	<td>
	<asp:Label runat="server" id="lblID" Text='<%# DataBinder.Eval(Container.DataItem, "employeeid")%>' />
	<asp:textbox runat="server" Font-size="10px"
  			Text='<%# DataBinder.Eval(Container.DataItem, "firstname")%>' />
	<asp:textbox runat="server" Font-size="10px"
  			Text='<%# DataBinder.Eval(Container.DataItem, "lastname")%>' />
	</td>
</template>

<template name="FooterTemplate">
   </tr>
   <tr>
    <hr>
   	<asp:LinkButton runat="server" id="btnPrev" Text="Prev" OnClick="PrevRecord" />&nbsp;&nbsp;
   	<asp:LinkButton runat="server" id="btnNext" Text="Next" OnClick="NextRecord" />
   </tr>
   </table>
</template>
</asp:datalist>

</form>
</body>
</html>