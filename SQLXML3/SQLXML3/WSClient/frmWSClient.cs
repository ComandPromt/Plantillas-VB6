using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Xml;
using System.Text;
using WSWeb;

namespace WSClient
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class frmWSClient : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button cmdExecute;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtParm1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtParm2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtParm3;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabDataSet;
		private System.Windows.Forms.TabPage tabXML;
		private System.Windows.Forms.TextBox txtXMLData;
		private System.Windows.Forms.DataGrid dataGrid1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmWSClient()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.cmdExecute = new System.Windows.Forms.Button();
			this.txtParm1 = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtParm2 = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.txtParm3 = new System.Windows.Forms.TextBox();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabDataSet = new System.Windows.Forms.TabPage();
			this.tabXML = new System.Windows.Forms.TabPage();
			this.txtXMLData = new System.Windows.Forms.TextBox();
			this.dataGrid1 = new System.Windows.Forms.DataGrid();
			this.tabControl1.SuspendLayout();
			this.tabDataSet.SuspendLayout();
			this.tabXML.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// cmdExecute
			// 
			this.cmdExecute.Location = new System.Drawing.Point(800, 24);
			this.cmdExecute.Name = "cmdExecute";
			this.cmdExecute.Size = new System.Drawing.Size(136, 23);
			this.cmdExecute.TabIndex = 0;
			this.cmdExecute.Text = "Execute";
			this.cmdExecute.Click += new System.EventHandler(this.cmdExecute_Click);
			// 
			// txtParm1
			// 
			this.txtParm1.Location = new System.Drawing.Point(80, 32);
			this.txtParm1.Name = "txtParm1";
			this.txtParm1.Size = new System.Drawing.Size(120, 20);
			this.txtParm1.TabIndex = 2;
			this.txtParm1.Text = "ALFKI";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 32);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(64, 23);
			this.label1.TabIndex = 3;
			this.label1.Text = "Parmater 1:";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(216, 32);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(64, 23);
			this.label2.TabIndex = 6;
			this.label2.Text = "Parmater 2:";
			// 
			// txtParm2
			// 
			this.txtParm2.Location = new System.Drawing.Point(280, 32);
			this.txtParm2.Name = "txtParm2";
			this.txtParm2.Size = new System.Drawing.Size(120, 20);
			this.txtParm2.TabIndex = 5;
			this.txtParm2.Text = "Beverages";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(416, 32);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(64, 23);
			this.label3.TabIndex = 8;
			this.label3.Text = "Parmater 3:";
			// 
			// txtParm3
			// 
			this.txtParm3.Location = new System.Drawing.Point(480, 32);
			this.txtParm3.Name = "txtParm3";
			this.txtParm3.Size = new System.Drawing.Size(120, 20);
			this.txtParm3.TabIndex = 7;
			this.txtParm3.Text = "1999";
			// 
			// tabControl1
			// 
			this.tabControl1.Controls.AddRange(new System.Windows.Forms.Control[] {
																					  this.tabDataSet,
																					  this.tabXML});
			this.tabControl1.Location = new System.Drawing.Point(8, 56);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(936, 528);
			this.tabControl1.TabIndex = 9;
			// 
			// tabDataSet
			// 
			this.tabDataSet.Controls.AddRange(new System.Windows.Forms.Control[] {
																					 this.dataGrid1});
			this.tabDataSet.Location = new System.Drawing.Point(4, 22);
			this.tabDataSet.Name = "tabDataSet";
			this.tabDataSet.Size = new System.Drawing.Size(928, 502);
			this.tabDataSet.TabIndex = 0;
			this.tabDataSet.Text = "DataSet";
			// 
			// tabXML
			// 
			this.tabXML.Controls.AddRange(new System.Windows.Forms.Control[] {
																				 this.txtXMLData});
			this.tabXML.Location = new System.Drawing.Point(4, 22);
			this.tabXML.Name = "tabXML";
			this.tabXML.Size = new System.Drawing.Size(928, 502);
			this.tabXML.TabIndex = 1;
			this.tabXML.Text = "XML";
			// 
			// txtXMLData
			// 
			this.txtXMLData.AcceptsReturn = true;
			this.txtXMLData.AcceptsTab = true;
			this.txtXMLData.Location = new System.Drawing.Point(8, 8);
			this.txtXMLData.Multiline = true;
			this.txtXMLData.Name = "txtXMLData";
			this.txtXMLData.ReadOnly = true;
			this.txtXMLData.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtXMLData.Size = new System.Drawing.Size(912, 488);
			this.txtXMLData.TabIndex = 6;
			this.txtXMLData.Text = "";
			// 
			// dataGrid1
			// 
			this.dataGrid1.DataMember = "";
			this.dataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dataGrid1.Location = new System.Drawing.Point(8, 8);
			this.dataGrid1.Name = "dataGrid1";
			this.dataGrid1.Size = new System.Drawing.Size(912, 488);
			this.dataGrid1.TabIndex = 0;
			// 
			// frmWSClient
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(952, 597);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.tabControl1,
																		  this.label3,
																		  this.txtParm3,
																		  this.label2,
																		  this.txtParm2,
																		  this.label1,
																		  this.txtParm1,
																		  this.cmdExecute});
			this.Name = "frmWSClient";
			this.Text = "Web Service Client";
			this.tabControl1.ResumeLayout(false);
			this.tabDataSet.ResumeLayout(false);
			this.tabXML.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new frmWSClient());
		}

		private void cmdExecute_Click(object sender, System.EventArgs e)
		{
			localhost.procedures oWSProcs = null;
			//int nReturnValue;
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
				
				// call stored proc number 2
				sOutput.Append("OrderByCategory StoredProcedure Results -----------");
				oXmlResult = WSLib.GetXmlFromObjectArray(oWSProcs.SalesByCategory(txtParm2.Text, txtParm3.Text));			
				ds = WSLib.FormatDataToBuffer("SalesByCategory", oXmlResult, sOutput);

				// call our template
				sOutput.Append("GetAllCustomers Template Results -----------");
				oXmlResult = WSLib.GetXmlFromObjectArray(oWSProcs.GetAllCustomers());			
				ds = WSLib.FormatDataToBuffer("GetAllCustomers", oXmlResult, sOutput);
				// show this one in the grid for fun
				dataGrid1.DataSource = ds;

				// call our UDF
				sOutput.Append("GetCustomerContactView UDF Results -----------");
				oXmlResult = WSLib.GetXmlFromObjectArray(oWSProcs.GetCustomerContactView(txtParm1.Text));			
				ds = WSLib.FormatDataToBuffer("GetCustomerContactView", oXmlResult, sOutput);

				// send it to screen
                txtXMLData.Text = sOutput.ToString();
				
			}
			catch (Exception err)
			{
				System.Windows.Forms.MessageBox.Show("Error Retrieving Results" + err.Message);
			}

		}

		

		
	}
}
