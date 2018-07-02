using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace InheritTextBoxClientCS
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private InheritTextBoxControlCS.InheritTextBoxControlCS inheritTextBoxControlCS1;
		private System.Windows.Forms.Label label1;
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
			this.inheritTextBoxControlCS1 = new InheritTextBoxControlCS.InheritTextBoxControlCS();
			this.label1 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// inheritTextBoxControlCS1
			// 
			this.inheritTextBoxControlCS1.BackColor = System.Drawing.Color.LightPink;
			this.inheritTextBoxControlCS1.Location = new System.Drawing.Point(40, 48);
			this.inheritTextBoxControlCS1.Name = "inheritTextBoxControlCS1";
			this.inheritTextBoxControlCS1.Size = new System.Drawing.Size(216, 20);
			this.inheritTextBoxControlCS1.TabIndex = 0;
			this.inheritTextBoxControlCS1.Text = "";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(40, 16);
			this.label1.Name = "label1";
			this.label1.TabIndex = 1;
			this.label1.Text = "E-mail address:";
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(336, 133);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.label1,
																		  this.inheritTextBoxControlCS1});
			this.Name = "Form1";
			this.Text = "Rolling Thunder InheritTextBox Client C#";
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
	}
}
