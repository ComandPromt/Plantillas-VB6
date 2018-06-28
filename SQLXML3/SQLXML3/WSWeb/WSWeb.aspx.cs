using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.IO;
using System.Text;
using System.Xml;
using System.Data;

namespace WSWeb
{
	/// <summary>
	/// Summary description for WebForm1.
	/// </summary>
	public class WSWebForm : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.Button cmdExecute;
		protected System.Web.UI.WebControls.Label Label1;
		protected System.Web.UI.WebControls.TextBox txtParm1;
		protected System.Web.UI.WebControls.Label Label4;
		protected System.Web.UI.WebControls.DataGrid DataGrid2;
		protected System.Web.UI.WebControls.Label Label5;
		protected System.Web.UI.WebControls.Label Label6;
		protected System.Web.UI.WebControls.DataGrid DataGrid1;
	
		private void Page_Load(object sender, System.EventArgs e)
		{
			localhost.procedures oWSProcs = null;
			int nReturnValue;
			XmlElement[] oXmlResult = null; 
			DataSet ds = null;
			StringBuilder sOutput = new StringBuilder();

			try
			{
				// instantiate our web service
				oWSProcs = new localhost.procedures();
				
				// call stored proc number 1
				sOutput.Append("CustOrderHist StoredProcedure Results -----------");
				oXmlResult = WSLib.GetXmlFromObjectArray(oWSProcs.CustOrderHist(txtParm1.Text));			
				ds = WSLib.FormatDataToBuffer("CustOrderHist", oXmlResult, sOutput);
				DataGrid1.DataSource = ds;
				DataGrid1.DataBind();

				// call our template
				sOutput.Append("GetAllCustomers Template Results -----------");
				oXmlResult = WSLib.GetXmlFromObjectArray(oWSProcs.GetAllCustomers());			
				ds = WSLib.FormatDataToBuffer("GetAllCustomers", oXmlResult, sOutput);
				DataGrid2.DataSource = ds;
				DataGrid2.DataBind();

			}
			catch (Exception err)
			{
				ErrorPage = "Error Retrieving Results" + err.Message;
			}
		}

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: This call is required by the ASP.NET Web Form Designer.
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{    
			this.cmdExecute.Click += new System.EventHandler(this.cmdExecute_Click);
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion

		private void cmdExecute_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}
