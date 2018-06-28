using System;
using System.Drawing;
using System.Collections;
using System.Windows.Forms;
using System.Data;
using Microsoft.Win32;
using System.Collections.Specialized;


namespace OpenDlg
{
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnOpen;
		private System.Windows.Forms.DataGrid grid;
		private System.Windows.Forms.Button btnOpenPlaces;
		private System.Windows.Forms.OpenFileDialog openFileDialog;

		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}


		public Form1()
		{
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}


		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.btnOpen = new System.Windows.Forms.Button();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.grid = new System.Windows.Forms.DataGrid();
			this.btnOpenPlaces = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.grid)).BeginInit();
			this.SuspendLayout();
			// 
			// btnOpen
			// 
			this.btnOpen.Location = new System.Drawing.Point(280, 8);
			this.btnOpen.Name = "btnOpen";
			this.btnOpen.Size = new System.Drawing.Size(104, 23);
			this.btnOpen.TabIndex = 0;
			this.btnOpen.Text = "Open...";
			this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
			// 
			// grid
			// 
			this.grid.DataMember = "";
			this.grid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.grid.Location = new System.Drawing.Point(8, 8);
			this.grid.Name = "grid";
			this.grid.ReadOnly = true;
			this.grid.Size = new System.Drawing.Size(256, 152);
			this.grid.TabIndex = 1;
			// 
			// btnOpenPlaces
			// 
			this.btnOpenPlaces.Location = new System.Drawing.Point(280, 40);
			this.btnOpenPlaces.Name = "btnOpenPlaces";
			this.btnOpenPlaces.Size = new System.Drawing.Size(104, 23);
			this.btnOpenPlaces.TabIndex = 2;
			this.btnOpenPlaces.Text = "Open Places...";
			this.btnOpenPlaces.Click += new System.EventHandler(this.btnOpenPlaces_Click);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(394, 167);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.btnOpenPlaces,
																		  this.grid,
																		  this.btnOpen});
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.Name = "Form1";
			this.Text = "PlacesBar Controller";
			this.Load += new System.EventHandler(this.Form1_Load);
			((System.ComponentModel.ISupportInitialize)(this.grid)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion



		private void btnOpen_Click(object sender, System.EventArgs e)
		{
			openFileDialog.InitialDirectory = @"c:\";
			openFileDialog.Filter = "Bitmap|*.bmp|All|*.*";
			openFileDialog.ShowDialog();
		}

		private const string Key_PlacesBar = @"Software\Microsoft\Windows\CurrentVersion\Policies\ComDlg32\PlacesBar";
		private void Form1_Load(object sender, System.EventArgs e)
		{
			FillTheGrid(Key_PlacesBar);
		}

	
		private void FillTheGrid(string key)
		{
			RegistryKey placesBarRoot = Registry.CurrentUser.OpenSubKey(key);
			if (placesBarRoot == null)
				return;
			string[] valuesOfKey = placesBarRoot.GetValueNames();
			
			DataTable dt = new DataTable(); 
			dt.Columns.Add("Place", typeof(string));
			dt.Columns.Add("Location", typeof(string));

			for(int i=0; i<valuesOfKey.Length; i++)
			{
				DataRow row = dt.NewRow();
				row["Place"] = valuesOfKey[i].ToString();
				row["Location"] = placesBarRoot.GetValue(valuesOfKey[i].ToString());
				dt.Rows.Add(row); 
			}
			grid.DataSource = dt;
			placesBarRoot.Close();
		}

		private void btnOpenPlaces_Click(object sender, System.EventArgs e)
		{
			OpenDialogPlaces o = new OpenDialogPlaces();
			o.Places.Add(@"c:\");
			o.Places.Add(17);
			o.Places.Add(5);
			o.Places.Add(@"c:\My Articles");
			o.Places.Add(6);
			o.Init();
			o.OpenDialog.ShowDialog();
			o.Reset();
		}
	}
}
