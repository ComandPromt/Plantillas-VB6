using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace UserControlClientCS
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private UserControlClientCS.UserControl1CS userControl1CS1;
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
			this.userControl1CS1 = new UserControlClientCS.UserControl1CS();
			this.SuspendLayout();
			// 
			// userControl1CS1
			// 
			this.userControl1CS1.BothTextBoxesBackColor = System.Drawing.SystemColors.Window;
			this.userControl1CS1.Location = new System.Drawing.Point(16, 24);
			this.userControl1CS1.Name = "userControl1CS1";
			this.userControl1CS1.Size = new System.Drawing.Size(256, 208);
			this.userControl1CS1.TabIndex = 0;
			this.userControl1CS1.CancelClicked += new UserControlClientCS.UserControl1CS.CancelClickedHandler(this.userControl1CS1_CancelClicked);
			this.userControl1CS1.OkClicked += new UserControlClientCS.UserControl1CS.OkClickedHandler(this.userControl1CS1_OkClicked);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(296, 261);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.userControl1CS1});
			this.Name = "Form1";
			this.Text = "Rolling Thunder UserControl Client C#";
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

		private void userControl1CS1_CancelClicked()
		{
			MessageBox.Show ("User clicked Cancel") ;
		}

		private void userControl1CS1_OkClicked(string UserID, string Password)
		{
			MessageBox.Show ("User clicked OK, UserID = " + UserID + ", Password = " + Password) ;
		}
	}
}
