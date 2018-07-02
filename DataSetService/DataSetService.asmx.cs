using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Web;
using System.Web.Services;
using System.Xml;

[WebService(Namespace="http://example.org/dataset-service")]
public class DataSetService : System.Web.Services.WebService
{
	public DataSetService()
	{
		//CODEGEN: This call is required by the ASP.NET Web Services Designer
		InitializeComponent();
	}

	private System.Data.SqlClient.SqlDataAdapter sqlDataAdapter1;
	private System.Data.SqlClient.SqlConnection sqlConnection;
	private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
	private System.Data.SqlClient.SqlCommand sqlInsertCommand1;
	private System.Data.SqlClient.SqlCommand sqlUpdateCommand1;
	private System.Data.SqlClient.SqlCommand sqlDeleteCommand1;

	#region Component Designer generated code
	
	//Required by the Web Services Designer 
	private IContainer components = null;
			
	/// <summary>
	/// Required method for Designer support - do not modify
	/// the contents of this method with the code editor.
	/// </summary>
	private void InitializeComponent()
	{
		this.sqlDataAdapter1 = new System.Data.SqlClient.SqlDataAdapter();
		this.sqlConnection = new System.Data.SqlClient.SqlConnection();
		this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
		this.sqlInsertCommand1 = new System.Data.SqlClient.SqlCommand();
		this.sqlUpdateCommand1 = new System.Data.SqlClient.SqlCommand();
		this.sqlDeleteCommand1 = new System.Data.SqlClient.SqlCommand();
		// 
		// sqlDataAdapter1
		// 
		this.sqlDataAdapter1.DeleteCommand = this.sqlDeleteCommand1;
		this.sqlDataAdapter1.InsertCommand = this.sqlInsertCommand1;
		this.sqlDataAdapter1.SelectCommand = this.sqlSelectCommand1;
		this.sqlDataAdapter1.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
																								  new System.Data.Common.DataTableMapping("Table", "authors", new System.Data.Common.DataColumnMapping[] {
																																																			 new System.Data.Common.DataColumnMapping("au_id", "au_id"),
																																																			 new System.Data.Common.DataColumnMapping("au_lname", "au_lname"),
																																																			 new System.Data.Common.DataColumnMapping("au_fname", "au_fname"),
																																																			 new System.Data.Common.DataColumnMapping("phone", "phone"),
																																																			 new System.Data.Common.DataColumnMapping("address", "address"),
																																																			 new System.Data.Common.DataColumnMapping("city", "city"),
																																																			 new System.Data.Common.DataColumnMapping("state", "state"),
																																																			 new System.Data.Common.DataColumnMapping("zip", "zip"),
																																																			 new System.Data.Common.DataColumnMapping("contract", "contract")})});
		this.sqlDataAdapter1.UpdateCommand = this.sqlUpdateCommand1;
		// 
		// sqlConnection
		// 
		this.sqlConnection.ConnectionString = "data source=.;initial catalog=pubs;persist security info=False;user id=sa;worksta" +
			"tion id=GANDALF;packet size=4096";
		// 
		// sqlSelectCommand1
		// 
		this.sqlSelectCommand1.CommandText = "SELECT au_id, au_lname, au_fname, phone, address, city, state, zip, contract FROM" +
			" authors";
		this.sqlSelectCommand1.Connection = this.sqlConnection;
		// 
		// sqlInsertCommand1
		// 
		this.sqlInsertCommand1.CommandText = @"INSERT INTO authors(au_id, au_lname, au_fname, phone, address, city, state, zip, contract) VALUES (@au_id, @au_lname, @au_fname, @phone, @address, @city, @state, @zip, @contract); SELECT au_id, au_lname, au_fname, phone, address, city, state, zip, contract FROM authors WHERE (au_id = @au_id)";
		this.sqlInsertCommand1.Connection = this.sqlConnection;
		this.sqlInsertCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@au_id", System.Data.SqlDbType.VarChar, 11, "au_id"));
		this.sqlInsertCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@au_lname", System.Data.SqlDbType.VarChar, 40, "au_lname"));
		this.sqlInsertCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@au_fname", System.Data.SqlDbType.VarChar, 20, "au_fname"));
		this.sqlInsertCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@phone", System.Data.SqlDbType.VarChar, 12, "phone"));
		this.sqlInsertCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@address", System.Data.SqlDbType.VarChar, 40, "address"));
		this.sqlInsertCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@city", System.Data.SqlDbType.VarChar, 20, "city"));
		this.sqlInsertCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@state", System.Data.SqlDbType.VarChar, 2, "state"));
		this.sqlInsertCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@zip", System.Data.SqlDbType.VarChar, 5, "zip"));
		this.sqlInsertCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@contract", System.Data.SqlDbType.Bit, 1, "contract"));
		// 
		// sqlUpdateCommand1
		// 
		this.sqlUpdateCommand1.CommandText = @"UPDATE authors SET au_id = @au_id, au_lname = @au_lname, au_fname = @au_fname, phone = @phone, address = @address, city = @city, state = @state, zip = @zip, contract = @contract WHERE (au_id = @Original_au_id) AND (address = @Original_address OR @Original_address IS NULL AND address IS NULL) AND (au_fname = @Original_au_fname) AND (au_lname = @Original_au_lname) AND (city = @Original_city OR @Original_city IS NULL AND city IS NULL) AND (contract = @Original_contract) AND (phone = @Original_phone) AND (state = @Original_state OR @Original_state IS NULL AND state IS NULL) AND (zip = @Original_zip OR @Original_zip IS NULL AND zip IS NULL); SELECT au_id, au_lname, au_fname, phone, address, city, state, zip, contract FROM authors WHERE (au_id = @au_id)";
		this.sqlUpdateCommand1.Connection = this.sqlConnection;
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@au_id", System.Data.SqlDbType.VarChar, 11, "au_id"));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@au_lname", System.Data.SqlDbType.VarChar, 40, "au_lname"));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@au_fname", System.Data.SqlDbType.VarChar, 20, "au_fname"));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@phone", System.Data.SqlDbType.VarChar, 12, "phone"));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@address", System.Data.SqlDbType.VarChar, 40, "address"));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@city", System.Data.SqlDbType.VarChar, 20, "city"));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@state", System.Data.SqlDbType.VarChar, 2, "state"));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@zip", System.Data.SqlDbType.VarChar, 5, "zip"));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@contract", System.Data.SqlDbType.Bit, 1, "contract"));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_au_id", System.Data.SqlDbType.VarChar, 11, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "au_id", System.Data.DataRowVersion.Original, null));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_address", System.Data.SqlDbType.VarChar, 40, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "address", System.Data.DataRowVersion.Original, null));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_au_fname", System.Data.SqlDbType.VarChar, 20, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "au_fname", System.Data.DataRowVersion.Original, null));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_au_lname", System.Data.SqlDbType.VarChar, 40, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "au_lname", System.Data.DataRowVersion.Original, null));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_city", System.Data.SqlDbType.VarChar, 20, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "city", System.Data.DataRowVersion.Original, null));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_contract", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "contract", System.Data.DataRowVersion.Original, null));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_phone", System.Data.SqlDbType.VarChar, 12, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "phone", System.Data.DataRowVersion.Original, null));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_state", System.Data.SqlDbType.VarChar, 2, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "state", System.Data.DataRowVersion.Original, null));
		this.sqlUpdateCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_zip", System.Data.SqlDbType.VarChar, 5, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "zip", System.Data.DataRowVersion.Original, null));
		// 
		// sqlDeleteCommand1
		// 
		this.sqlDeleteCommand1.CommandText = @"DELETE FROM authors WHERE (au_id = @Original_au_id) AND (address = @Original_address OR @Original_address IS NULL AND address IS NULL) AND (au_fname = @Original_au_fname) AND (au_lname = @Original_au_lname) AND (city = @Original_city OR @Original_city IS NULL AND city IS NULL) AND (contract = @Original_contract) AND (phone = @Original_phone) AND (state = @Original_state OR @Original_state IS NULL AND state IS NULL) AND (zip = @Original_zip OR @Original_zip IS NULL AND zip IS NULL)";
		this.sqlDeleteCommand1.Connection = this.sqlConnection;
		this.sqlDeleteCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_au_id", System.Data.SqlDbType.VarChar, 11, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "au_id", System.Data.DataRowVersion.Original, null));
		this.sqlDeleteCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_address", System.Data.SqlDbType.VarChar, 40, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "address", System.Data.DataRowVersion.Original, null));
		this.sqlDeleteCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_au_fname", System.Data.SqlDbType.VarChar, 20, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "au_fname", System.Data.DataRowVersion.Original, null));
		this.sqlDeleteCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_au_lname", System.Data.SqlDbType.VarChar, 40, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "au_lname", System.Data.DataRowVersion.Original, null));
		this.sqlDeleteCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_city", System.Data.SqlDbType.VarChar, 20, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "city", System.Data.DataRowVersion.Original, null));
		this.sqlDeleteCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_contract", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "contract", System.Data.DataRowVersion.Original, null));
		this.sqlDeleteCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_phone", System.Data.SqlDbType.VarChar, 12, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "phone", System.Data.DataRowVersion.Original, null));
		this.sqlDeleteCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_state", System.Data.SqlDbType.VarChar, 2, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "state", System.Data.DataRowVersion.Original, null));
		this.sqlDeleteCommand1.Parameters.Add(new System.Data.SqlClient.SqlParameter("@Original_zip", System.Data.SqlDbType.VarChar, 5, System.Data.ParameterDirection.Input, false, ((System.Byte)(0)), ((System.Byte)(0)), "zip", System.Data.DataRowVersion.Original, null));

	}

	/// <summary>
	/// Clean up any resources being used.
	/// </summary>
	protected override void Dispose( bool disposing )
	{
		if(disposing && components != null)
		{
			components.Dispose();
		}
		base.Dispose(disposing);		
	}
	
	#endregion

	[WebMethod]
	public DataSet GetAuthors()
	{
		DataSet ds = new DataSet("authors");
		sqlDataAdapter1.Fill(ds);
		ds.Tables[0].TableName = "author";
		return ds;
	}

	[WebMethod]
	public AuthorSet GetAuthorsAsTypedDataSet()
	{
		AuthorSet aus = new AuthorSet();
		sqlDataAdapter1.Fill(aus);
		return aus;
	}

	[WebMethod]
	public XmlNode GetAuthorsAsXml()
	{
		AuthorSet aus = new AuthorSet();
		sqlDataAdapter1.Fill(aus);
		return new XmlDataDocument(aus);
	}
}

