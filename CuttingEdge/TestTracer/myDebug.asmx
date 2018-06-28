<%@ WebService language="C#" class="MyDebugTool" %>

using System;
using System.Web.Services;
using System.Data;
using System.IO;
using System.Data.SqlClient;


[WebService(Namespace="MsdnMag.CuttingEdge")]
public class MyDebugTool 
{
	[WebMethod]
	public DataSet GetInfo(string connString, string userKey)
	{
		SqlDataAdapter adapter = new SqlDataAdapter(
			"SELECT data FROM internalcache WHERE userkey='" + userKey + "'",
			connString);
		DataTable tmp = new DataTable();
		adapter.Fill(tmp);
		
		DataSet ds = new DataSet();
		StringReader reader = new StringReader(tmp.Rows[0]["data"].ToString());
		ds.ReadXml(reader);		
		ds.AcceptChanges();
		return ds;
	}
}