<%@ Control ClassName="MyDebugTool" Language="C#" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>


<script runat="server">
// *****************************************************************
// Run when the host page is unloading and caches info
private void Page_Unload(object sender, EventArgs e)
{
	GrabAndStoreContextInfo();
}  		
// *****************************************************************


// *****************************************************************
// Run when the host page is going to render 
private void Page_PreRender(object sender, EventArgs e)
{
	GrabAndStoreRequestInfo();
}  		
// *****************************************************************



// Key for the database
public string UserKey = "__MyDebugTool"; 	

// Connection string for the database
public string ConnString = "SERVER=localhost;DATABASE=MyTracer;UID=sa;";		

// Should load all possible info
public bool ShowAll = true;		

// Bind to the page's view state (necessary because viewstate is protected)
private StateBag m_boundViewState = null;
public void BindViewState(StateBag state)
{
	m_boundViewState = state;
}




// *****************************************************************
// Internals
private DataSet info = new DataSet("MyTracer");
// *****************************************************************


// *****************************************************************
// Collect all the information from the HTTP context
private void GrabAndStoreContextInfo()
{
	// Application
	LoadFromApplication(info);

	// Session
	LoadFromSession(info);
	
	// Cache
	LoadFromCache(info);

	// Make the DataSet publicly available
	PublishDataSet(info);
}
// *****************************************************************

// *****************************************************************
// Collect all the request information from the HTTP context
private void GrabAndStoreRequestInfo()
{
	// View state
	LoadFromViewState(info);

	// Controls
	LoadFromPageControls(info);

	// Form + QueryString
	LoadFromRequest(info);

	// Request Headers
	LoadFromRequestHeaders(info);

	// Server Variables
	LoadFromServerVariables(info);
	
	// Cookies
	LoadFromCookies(info);
}
// *****************************************************************


// *****************************************************************
// Load information from Cache
private void LoadFromCache(DataSet info)
{
	DataTable dtCache = CreateKeyValueDataTable("Cache");
	foreach(DictionaryEntry elem in Page.Cache)
	{
		if (ShowAll)
		{
			AddKeyValueItemToTable(dtCache, elem.Key.ToString(), DisplayFormat(elem.Value));
		}
		else
		{	
			string s = elem.Key.ToString();
			if (!s.StartsWith("ISAPIWorkerRequest") && !s.StartsWith("System"))
				AddKeyValueItemToTable(dtCache, elem.Key.ToString(), DisplayFormat(elem.Value)); 
		}
	}

	dtCache.AcceptChanges();
	info.Tables.Add(dtCache);
	return;
}
// *****************************************************************

// *****************************************************************
// Load information from page's controls
private void LoadFromPageControls(DataSet info)
{
	HtmlForm theForm = null;
	
	DataTable dtControls = CreateKeyValueDataTable("Controls");
	for (int i=0; i<Page.Controls.Count; i++)
	{
		if (Page.Controls[i] is HtmlForm)
			theForm = (HtmlForm) Page.Controls[i];
		AddKeyValueItemToTable(dtControls, Page.Controls[i].ToString(), 
			Page.Controls[i].ClientID);
	}
	dtControls.AcceptChanges();
	info.Tables.Add(dtControls);

	if (theForm == null)
		return;	
	
	DataTable dtFormControls = CreateKeyValueDataTable("FormControls");
	for (int i=0; i<theForm.Controls.Count; i++)
		AddKeyValueItemToTable(dtFormControls, theForm.Controls[i].ToString(), 
			theForm.Controls[i].ClientID);
	dtFormControls.AcceptChanges();
	info.Tables.Add(dtFormControls);

	return;
}
// *****************************************************************

// *****************************************************************
// Load information from Application
private void LoadFromApplication(DataSet info)
{
	DataTable dtApp = CreateKeyValueDataTable("Application");
	for (int i=0; i<Application.Count; i++)
		AddKeyValueItemToTable(dtApp, Application.Keys[i].ToString(), 
			DisplayFormat(Application[i]));

	dtApp.AcceptChanges();
	info.Tables.Add(dtApp);
	return;
}
// *****************************************************************

// *****************************************************************
// Load information from Session
private void LoadFromSession(DataSet info)
{
	DataTable dtSession = CreateKeyValueDataTable("Session");
	for (int i=0; i<Session.Count; i++)
		AddKeyValueItemToTable(dtSession, Session.Keys[i].ToString(), 
			DisplayFormat(Session[i]));

	dtSession.AcceptChanges();
	info.Tables.Add(dtSession);
	return;
}
// *****************************************************************

// *****************************************************************
// Load information from ViewState
private void LoadFromViewState(DataSet info)
{
	if (m_boundViewState == null)
		return;
		
	DataTable dtViewState = CreateKeyValueDataTable("ViewState");
	string [] keys = new string[m_boundViewState.Count]; 
	m_boundViewState.Keys.CopyTo(keys, 0);
	StateItem [] values = new StateItem[m_boundViewState.Count];
	m_boundViewState.Values.CopyTo(values, 0);
	
	for (int i=0; i<m_boundViewState.Count; i++)
	{
		AddKeyValueItemToTable(dtViewState, keys[i].ToString(), 
			DisplayFormat(values[i].Value));
	}

	dtViewState.AcceptChanges();
	info.Tables.Add(dtViewState);
	return;
}
// *****************************************************************


// *****************************************************************
// Load information from Request.Headers
private void LoadFromRequestHeaders(DataSet info)
{
	DataTable dtHeaders = CreateKeyValueDataTable("Headers");
	for (int i=0; i<Request.Headers.Count; i++)
		AddKeyValueItemToTable(dtHeaders, Request.Headers.Keys[i].ToString(), 
			DisplayFormat(Request.Headers[i]));

	dtHeaders.AcceptChanges();
	info.Tables.Add(dtHeaders);
	return;
}
// *****************************************************************

// *****************************************************************
// Load information from Request.ServerVariables
private void LoadFromServerVariables(DataSet info)
{
	DataTable dtServerVars = CreateKeyValueDataTable("ServerVariables");
	for (int i=0; i<Request.ServerVariables.Count; i++)
		AddKeyValueItemToTable(dtServerVars, Request.ServerVariables.Keys[i].ToString(), 
			Request.ServerVariables[i]);

	dtServerVars.AcceptChanges();
	info.Tables.Add(dtServerVars);
	return;
}
// *****************************************************************

// *****************************************************************
// Load information from Request.Cookies
private void LoadFromCookies(DataSet info)
{
	DataTable dtCookies = CreateKeyValueDataTable("Cookies");
	for (int i=0; i<Request.Cookies.Count; i++)
		AddKeyValueItemToTable(dtCookies, Request.Cookies.Keys[i].ToString(), 
			Request.Cookies[i].Value);

	dtCookies.AcceptChanges();
	info.Tables.Add(dtCookies);
	return;
}
// *****************************************************************

// *****************************************************************
// Load information from Request.Form  
private void LoadFromRequest(DataSet info)
{
	DataTable dtForm = CreateKeyValueDataTable("Form");
	for (int i=0; i<Request.Form.Count; i++)
	{
		string key = Request.Form.Keys[i].ToString();
		if (key == "__VIEWSTATE")
			AddKeyValueItemToTable(dtForm, key, "{ view state }"); 
		else
			AddKeyValueItemToTable(dtForm, key, Request.Form[i]); 
	}

	dtForm.AcceptChanges();
	info.Tables.Add(dtForm);
	return;
}
// *****************************************************************






// *****************************************************************
// Create the table with the expected Key/Value structure
private DataTable CreateKeyValueDataTable(string tableName)
{
	DataTable dt = new DataTable(tableName);
	dt.Columns.Add("Key", typeof(string));
	dt.Columns.Add("Value", typeof(string));
	return dt;
}
// *****************************************************************

// *****************************************************************
// Add a row to a key/value table
private void AddKeyValueItemToTable(DataTable dt, string key, string value)
{
	DataRow row = dt.NewRow();
	row["Key"] = key;
	row["Value"] = value;
	dt.Rows.Add(row);
}
// *****************************************************************

// *****************************************************************
// Serialize the overall DataSet to the database
private void PublishDataSet(DataSet info)
{
	// Serialize to a diffgram
	StringWriter writer = new StringWriter();
	info.WriteXml(writer); 

	// Assume a record with the specified UserKey is already in the db
	SqlConnection conn = new SqlConnection(ConnString);
	SqlCommand cmd = new SqlCommand("UPDATE InternalCache SET Data=@TheData WHERE UserKey=@TheUser", conn);
	cmd.Parameters.Add("@TheData", SqlDbType.Text).Value = writer.ToString();
	cmd.Parameters.Add("@TheUser", SqlDbType.VarChar).Value = UserKey;
	conn.Open();
	cmd.ExecuteNonQuery();
	conn.Close();
}
// *****************************************************************


// *****************************************************************
// Format strings and object references
private string DisplayFormat(object o)
{
	return (o is string || o.GetType().IsPrimitive
		?Convert.ToString(o) :"{ " + o.ToString() + " }");
}
// *****************************************************************
</script>

<table style="font-family:verdana;font-size:8pt;border:solid 1px;" 
	width="100%" bgcolor="cyan"><tr><td>
This page is traced using "<b>MyTracer</b>", courtesy of Cutting Edge
</td></tr></table>
	
