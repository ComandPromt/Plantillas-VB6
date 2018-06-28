using System;
using System.Windows.Forms;
using System.Runtime.InteropServices; 
using System.Collections;
using Microsoft.Win32;
using System.Reflection;


namespace OpenDlg
{

	// cannot inherit from OpenFileDialog (sealed)
	public class OpenDialogPlaces 
	{
		private const string Key_PlacesBar = @"Software\Microsoft\Windows\CurrentVersion\Policies\ComDlg32\PlacesBar";
		private RegistryKey m_fakeKey;
		private IntPtr m_overriddenKey;
		private OpenFileDialog m_openFileDialog;
		private ArrayList m_places;

		public OpenFileDialog OpenDialog
		{
			get {return m_openFileDialog;}
		}

		public ArrayList Places
		{
			get {return m_places;}
		}

		public OpenDialogPlaces()
		{
			m_places = new ArrayList();
			m_openFileDialog = new OpenFileDialog();				
		}

		public void Init()
		{
			SetupFakeRegistryTree();
		}

		public void Reset()
		{
			ResetRegistry(m_overriddenKey);

			// should delete the key
			//m_fakeKey.DeleteSubKeyTree("Dino"); 
		}

		private void SetupFakeRegistryTree()
		{
			m_fakeKey = Registry.CurrentUser.CreateSubKey("dino");
			m_overriddenKey = InitializeRegistry();

			// at this point, "dino" equals places key
			// write dynamic places here reading from Places

			for(int i=0; i<Places.Count; i++)
			{
				if(Places[i] != null)
				{
					RegistryKey reg = Registry.CurrentUser.CreateSubKey(Key_PlacesBar);
					reg.SetValue("Place" + i.ToString(), Places[i]);  
				}
			}
		}

		[DllImport("myregutil.dll")]
		private static extern IntPtr InitializeRegistry();
		[DllImport("myregutil.dll")]
		private static extern int ResetRegistry(IntPtr hKey);
	}
}
