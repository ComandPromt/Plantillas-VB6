using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace BlinkingLabelClientCS
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{

		private BlinkingLabelControlCS.BlinkingLabelControl blinkingLabelControl1;
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
			this.blinkingLabelControl1 = new BlinkingLabelControlCS.BlinkingLabelControl();
			this.SuspendLayout();
			// 
			// blinkingLabelControl1
			// 
			this.blinkingLabelControl1.BlinkInterval = 1;
			this.blinkingLabelControl1.BlinkOffColor = System.Drawing.SystemColors.Control;
			this.blinkingLabelControl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.blinkingLabelControl1.Location = new System.Drawing.Point(32, 24);
			this.blinkingLabelControl1.Name = "blinkingLabelControl1";
			this.blinkingLabelControl1.Size = new System.Drawing.Size(208, 32);
			this.blinkingLabelControl1.TabIndex = 0;
			this.blinkingLabelControl1.Text = "blinkingLabelControl1";
			this.blinkingLabelControl1.BlinkStateChanged += new BlinkingLabelControlCS.BlinkingLabelControl.BlinkStateChangedHandler(this.blinkingLabelControl1_BlinkStateChanged);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(312, 165);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.blinkingLabelControl1});
			this.Name = "Form1";
			this.Text = "Rolling Thunder Blinking Label Demo C#";
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


		[System.Runtime.InteropServices.DllImport("user32.dll")]
		public static extern bool MessageBeep (uint uType);

		private void blinkingLabelControl1_BlinkStateChanged(bool UseBlinkColor)
		{
			MessageBeep (0) ;
		}
	}
}
