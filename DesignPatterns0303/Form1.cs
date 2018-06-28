using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace FactoryPattern
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnFPTest;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.ComboBox cmbPart;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form1()
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
			this.cmbPart = new System.Windows.Forms.ComboBox();
			this.btnFPTest = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.SuspendLayout();
			// 
			// cmbPart
			// 
			this.cmbPart.Location = new System.Drawing.Point(96, 40);
			this.cmbPart.Name = "cmbPart";
			this.cmbPart.Size = new System.Drawing.Size(136, 21);
			this.cmbPart.TabIndex = 0;
			// 
			// btnFPTest
			// 
			this.btnFPTest.BackColor = System.Drawing.Color.Honeydew;
			this.btnFPTest.Location = new System.Drawing.Point(96, 112);
			this.btnFPTest.Name = "btnFPTest";
			this.btnFPTest.Size = new System.Drawing.Size(136, 32);
			this.btnFPTest.TabIndex = 2;
			this.btnFPTest.Text = "&Dynamic Factory Test";
			this.btnFPTest.Click += new System.EventHandler(this.btnFPTest_Click);
			// 
			// btnClose
			// 
			this.btnClose.BackColor = System.Drawing.Color.Honeydew;
			this.btnClose.Location = new System.Drawing.Point(96, 152);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(136, 32);
			this.btnClose.TabIndex = 3;
			this.btnClose.Text = "&Exit";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(24, 16);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(288, 200);
			this.groupBox1.TabIndex = 4;
			this.groupBox1.TabStop = false;
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.BackColor = System.Drawing.Color.MediumAquamarine;
			this.ClientSize = new System.Drawing.Size(336, 246);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.btnClose,
																		  this.btnFPTest,
																		  this.cmbPart,
																		  this.groupBox1});
			this.Name = "Form1";
			this.Text = "Concrete Factory Pattern";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

		private void btnFPTest_Click(object sender, System.EventArgs e)
		{			
			InventoryMgr InvMgr = new InventoryMgr();
			cmbPart.Items.GetEnumerator();
			
			foreach(object marg in cmbPart.Items)
			{
				switch((string)marg) 
				{
					case "Monitors":                 
					    InvMgr.ReplenishInventory(enmInvParts.Monitors);                 
						break;
					case "Keyboards":
						InvMgr.ReplenishInventory(enmInvParts.Keyboards);
						break;
					case "MousePads":                    
                        InvMgr.ReplenishInventory(enmInvParts.Keyboards);          
                        break;
					default:
						break;
				}				                
			}
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			cmbPart.Items.Add((object)"Monitors");
			cmbPart.Items.Add((object)"Keyboards");
			cmbPart.Items.Add((object)"MousePads");			
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}	
	}
}
