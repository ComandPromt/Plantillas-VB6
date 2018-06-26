////////////////////////////////////////////////////////////////
// MSDN Magazine -- March 2003
// If this code works, it was written by Paul DiLascia.
// If not, I don't know who wrote it.
// Compiles with Visual Studio .NET on Windows XP. Tab size=3.
//
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;

namespace WinApp
{
	public class Form1 : System.Windows.Forms.Form
	{
		public Form1()
		{
			BackColor = SystemColors.Window;
		}

		protected override void OnPaint(PaintEventArgs e)
		{
			// get property info for all SystemIcons static properties
			PropertyInfo[] props = typeof(SystemIcons).GetProperties(
				BindingFlags.Public|BindingFlags.Static); 

			Graphics g = e.Graphics;
			Font font = new Font("Verdana", 12, FontStyle.Bold);
			SolidBrush brush = new SolidBrush(Color.Black);
			int y = 0;

			// Display each icon. Use reflection to get all the static
			// members of SysIcons--cool!
			//
			foreach (PropertyInfo p in props) {
				Object obj = p.GetValue(null, null);
				if (obj.GetType()==typeof(Icon)) {
					Icon icon = (Icon)obj;
					g.DrawIcon(icon, 0, y);
					g.DrawString(String.Format("SystemIcons.{0}",p.Name),
						font, brush, icon.Width+2, y);
					y += icon.Height;
				}
			}
		}

		protected override Size DefaultSize
		{
			get { return new Size(300,350); }
		}

		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}
	}
}
