<%@ webhandler language="C#" class="PubLogoHandler" %>

using System;
using System.Web;
using System.Data;
using System.Data.SqlClient;

public class PubLogoHandler : IHttpHandler
{
	public bool IsReusable
	{
		get { return true; }
	}
	
	public void ProcessRequest(HttpContext ctx)
	{
		string pubID = ctx.Request.QueryString["id"];
		
		SqlConnection con = new SqlConnection("server=localhost;integrated security=SSPI;database=pubs");
		SqlCommand cmd = new SqlCommand("SELECT logo FROM pub_info WHERE pub_id = @pubID", con);
		cmd.CommandType = CommandType.Text;
		cmd.Parameters.Add("@pubID", pubID);

		con.Open();
		byte[] logo = (byte[])cmd.ExecuteScalar();
		con.Close();

		ctx.Response.ContentType = "image/gif";
		ctx.Response.OutputStream.Write(logo, 0, logo.Length);
 	}
}