using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace InheritTextBoxControlCS
{
	/// <summary>
	/// Summary description for InheritTextBoxControlCS.
	/// </summary>
	public class InheritTextBoxControlCS : System.Windows.Forms.TextBox
	{
		public InheritTextBoxControlCS()
		{
		}

		// Check to see if string in text box appears to be a valid e-mail address.
		// In this case, that means that it contains at least one @ sign and
		// at least one period.

		protected override void OnTextChanged(System.EventArgs e)
		{
			// Pass call to base class
			
			base.OnTextChanged(e) ;

			// Perform our checking logic using inherited property Text, and
			// set value of inherited property BackColor accordingly

			if (Text.IndexOf("@") != -1 && Text.IndexOf(".") != -1)
			{
				BackColor = System.Drawing.Color.LightGreen ;
			}
			else
			{
				BackColor = System.Drawing.Color.LightPink ;
			}
		}
	}
}
