using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace UserControlClientCS
{
	/// <summary>
	/// Summary description for UserControl1CS.
	/// </summary>
	public class UserControl1CS : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox textBox1;
		private System.Windows.Forms.TextBox textBox2;
		private System.Windows.Forms.ErrorProvider errorProvider1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button2;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public UserControl1CS()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitForm call

		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.textBox2 = new System.Windows.Forms.TextBox();
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.button2 = new System.Windows.Forms.Button();
			this.errorProvider1 = new System.Windows.Forms.ErrorProvider();
			this.button1 = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// textBox2
			// 
			this.textBox2.Location = new System.Drawing.Point(16, 104);
			this.textBox2.Name = "textBox2";
			this.textBox2.PasswordChar = '*';
			this.textBox2.Size = new System.Drawing.Size(184, 20);
			this.textBox2.TabIndex = 3;
			this.textBox2.Text = "";
			// 
			// textBox1
			// 
			this.textBox1.Location = new System.Drawing.Point(16, 40);
			this.textBox1.Name = "textBox1";
			this.textBox1.Size = new System.Drawing.Size(184, 20);
			this.textBox1.TabIndex = 2;
			this.textBox1.Text = "";
			// 
			// button2
			// 
			this.button2.Location = new System.Drawing.Point(128, 168);
			this.button2.Name = "button2";
			this.button2.TabIndex = 5;
			this.button2.Text = "Cancel";
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(16, 168);
			this.button1.Name = "button1";
			this.button1.TabIndex = 4;
			this.button1.Text = "OK";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 16);
			this.label1.Name = "label1";
			this.label1.TabIndex = 0;
			this.label1.Text = "User ID:";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(16, 88);
			this.label2.Name = "label2";
			this.label2.TabIndex = 1;
			this.label2.Text = "Password:";
			// 
			// UserControl1CS
			// 
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.button2,
																		  this.button1,
																		  this.textBox2,
																		  this.textBox1,
																		  this.label2,
																		  this.label1});
			this.Name = "UserControl1CS";
			this.Size = new System.Drawing.Size(256, 208);
			this.ResumeLayout(false);

		}
		#endregion

		// User clicked OK

		private void button1_Click(object sender, System.EventArgs e)
		{
			 bool bFieldsValid = true ;

			// Check to make sure that required fields are filled in.
			// Set error provider control to signal errors to the user if they're not.

			if (textBox1.Text.Length == 0)
			{
	            errorProvider1.SetError(textBox1, "A UserID is required") ;
				bFieldsValid = false ;
			}
			else
			{
				errorProvider1.SetError(textBox1, "") ;
			}

			if (textBox2.Text.Length == 0)
			{
				errorProvider1.SetError(textBox2, "A Password is required") ;
				bFieldsValid = false ;
			}
			else
			{
				errorProvider1.SetError(textBox2, "") ;
			}

			// Fire event to container if they are. 

			if (bFieldsValid == true) 
			{
				OkClicked(textBox1.Text, textBox2.Text) ;
			}
		}

		//User clicked Cancel. Fire event to its container.

		private void button2_Click(object sender, System.EventArgs e)
		{
			CancelClicked() ;
		}

		// Declare the events that this control will fire to its container. There's
		// one for the OK button and one for the Cancel button.

		public delegate void OkClickedHandler (string UserID, string Password) ;
		public event OkClickedHandler OkClicked ;

		public delegate void CancelClickedHandler () ;
		public event CancelClickedHandler CancelClicked ;

		System.Drawing.Pen m_Pen  = new System.Drawing.Pen (Color.Black, 3) ;

		protected override void OnPaint (System.Windows.Forms.PaintEventArgs pevent)
		{
			// forward call to base class

			base.OnPaint (pevent) ;

			// add our own customization

	        pevent.Graphics.DrawRectangle(m_Pen, 0, 0, this.Bounds.Width - 1, this.Bounds.Height - 1) ;
		}

		// Custom property, the background color used by both 
		// constituent text boxes

		private System.Drawing.Color m_BothTextBoxesBackColor  = Color.FromKnownColor(KnownColor.Window) ;
		
		public System.Drawing.Color BothTextBoxesBackColor
		{
			get
			{
				return m_BothTextBoxesBackColor;
			}
			set
			{
				m_BothTextBoxesBackColor = value ;
			}
		}
	}
}
