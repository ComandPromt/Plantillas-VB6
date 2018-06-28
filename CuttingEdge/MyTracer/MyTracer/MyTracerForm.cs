using System;
using System.Drawing;
using System.Collections;
using System.Windows.Forms;
using System.Data;
using System.Net;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.UI;
using AxSHDocVw;

namespace MsdnMag.CuttingEdge
{
	public class MyTracerForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.TabPage tabBrowser;
		private AxSHDocVw.AxWebBrowser webBrowser;
		private System.Windows.Forms.StatusBar appStatusBar;
		private System.Windows.Forms.TabControl tabs;

		private string m_currentPage = "";
		private const string APP_TITLE = "My Tracer";
		private System.Windows.Forms.TabPage tabCache;
		private System.Windows.Forms.Button buttonGo;
		private System.Windows.Forms.TextBox addressBar;
		private System.Windows.Forms.DataGrid cacheGrid;
		private const string APP_DEFAULTURL = "about:Type%20a%20URL";
		private System.Windows.Forms.TabPage tabApplication;
		private System.Windows.Forms.TabPage tabHeaders;
		private System.Windows.Forms.TabPage tabRequest;
		private System.Windows.Forms.TabPage tabServerVars;
		private System.Windows.Forms.TabPage tabCookies;
		private System.Windows.Forms.TabPage tabPageControls;
		private System.Windows.Forms.DataGrid serverGrid;
		private System.Windows.Forms.DataGrid cookiesGrid;
		private System.Windows.Forms.DataGrid headersGrid;
		private System.Windows.Forms.DataGrid appGrid;
		private System.Windows.Forms.DataGrid formGrid;
		private System.Windows.Forms.TabPage tabSession;
		private System.Windows.Forms.TabPage tabViewState;
		private System.Windows.Forms.DataGrid sessionGrid;
		private System.Windows.Forms.DataGrid viewstateGrid;
		private System.Windows.Forms.DataGrid controlsGrid;
		private System.Windows.Forms.DataGrid formControlsGrid;

		private const string CONNSTRING = "SERVER=localhost;DATABASE=MyTracer;UID=sa;";


		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MyTracerForm));
			this.tabs = new System.Windows.Forms.TabControl();
			this.tabBrowser = new System.Windows.Forms.TabPage();
			this.webBrowser = new AxSHDocVw.AxWebBrowser();
			this.tabApplication = new System.Windows.Forms.TabPage();
			this.appGrid = new System.Windows.Forms.DataGrid();
			this.tabCache = new System.Windows.Forms.TabPage();
			this.cacheGrid = new System.Windows.Forms.DataGrid();
			this.tabSession = new System.Windows.Forms.TabPage();
			this.sessionGrid = new System.Windows.Forms.DataGrid();
			this.tabViewState = new System.Windows.Forms.TabPage();
			this.viewstateGrid = new System.Windows.Forms.DataGrid();
			this.tabHeaders = new System.Windows.Forms.TabPage();
			this.headersGrid = new System.Windows.Forms.DataGrid();
			this.tabRequest = new System.Windows.Forms.TabPage();
			this.formGrid = new System.Windows.Forms.DataGrid();
			this.tabCookies = new System.Windows.Forms.TabPage();
			this.cookiesGrid = new System.Windows.Forms.DataGrid();
			this.tabServerVars = new System.Windows.Forms.TabPage();
			this.serverGrid = new System.Windows.Forms.DataGrid();
			this.tabPageControls = new System.Windows.Forms.TabPage();
			this.formControlsGrid = new System.Windows.Forms.DataGrid();
			this.controlsGrid = new System.Windows.Forms.DataGrid();
			this.appStatusBar = new System.Windows.Forms.StatusBar();
			this.buttonGo = new System.Windows.Forms.Button();
			this.addressBar = new System.Windows.Forms.TextBox();
			this.tabs.SuspendLayout();
			this.tabBrowser.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.webBrowser)).BeginInit();
			this.tabApplication.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.appGrid)).BeginInit();
			this.tabCache.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cacheGrid)).BeginInit();
			this.tabSession.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.sessionGrid)).BeginInit();
			this.tabViewState.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.viewstateGrid)).BeginInit();
			this.tabHeaders.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.headersGrid)).BeginInit();
			this.tabRequest.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.formGrid)).BeginInit();
			this.tabCookies.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cookiesGrid)).BeginInit();
			this.tabServerVars.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.serverGrid)).BeginInit();
			this.tabPageControls.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.formControlsGrid)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.controlsGrid)).BeginInit();
			this.SuspendLayout();
			// 
			// tabs
			// 
			this.tabs.Alignment = System.Windows.Forms.TabAlignment.Bottom;
			this.tabs.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.tabs.Controls.AddRange(new System.Windows.Forms.Control[] {
																			   this.tabBrowser,
																			   this.tabApplication,
																			   this.tabCache,
																			   this.tabSession,
																			   this.tabViewState,
																			   this.tabHeaders,
																			   this.tabRequest,
																			   this.tabCookies,
																			   this.tabServerVars,
																			   this.tabPageControls});
			this.tabs.Location = new System.Drawing.Point(8, 40);
			this.tabs.Multiline = true;
			this.tabs.Name = "tabs";
			this.tabs.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.tabs.SelectedIndex = 0;
			this.tabs.Size = new System.Drawing.Size(706, 464);
			this.tabs.TabIndex = 0;
			// 
			// tabBrowser
			// 
			this.tabBrowser.Controls.AddRange(new System.Windows.Forms.Control[] {
																					 this.webBrowser});
			this.tabBrowser.Location = new System.Drawing.Point(4, 4);
			this.tabBrowser.Name = "tabBrowser";
			this.tabBrowser.Size = new System.Drawing.Size(698, 438);
			this.tabBrowser.TabIndex = 0;
			this.tabBrowser.Text = "Browser";
			// 
			// webBrowser
			// 
			this.webBrowser.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.webBrowser.ContainingControl = this;
			this.webBrowser.Enabled = true;
			this.webBrowser.Location = new System.Drawing.Point(8, 8);
			this.webBrowser.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("webBrowser.OcxState")));
			this.webBrowser.Size = new System.Drawing.Size(682, 424);
			this.webBrowser.TabIndex = 0;
			// 
			// tabApplication
			// 
			this.tabApplication.Controls.AddRange(new System.Windows.Forms.Control[] {
																						 this.appGrid});
			this.tabApplication.Location = new System.Drawing.Point(4, 4);
			this.tabApplication.Name = "tabApplication";
			this.tabApplication.Size = new System.Drawing.Size(698, 438);
			this.tabApplication.TabIndex = 2;
			this.tabApplication.Text = "Application";
			// 
			// appGrid
			// 
			this.appGrid.AllowNavigation = false;
			this.appGrid.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.appGrid.CaptionText = "Application";
			this.appGrid.DataMember = "";
			this.appGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.appGrid.Location = new System.Drawing.Point(8, 7);
			this.appGrid.Name = "appGrid";
			this.appGrid.ReadOnly = true;
			this.appGrid.Size = new System.Drawing.Size(682, 424);
			this.appGrid.TabIndex = 3;
			// 
			// tabCache
			// 
			this.tabCache.Controls.AddRange(new System.Windows.Forms.Control[] {
																				   this.cacheGrid});
			this.tabCache.Location = new System.Drawing.Point(4, 4);
			this.tabCache.Name = "tabCache";
			this.tabCache.Size = new System.Drawing.Size(698, 438);
			this.tabCache.TabIndex = 1;
			this.tabCache.Text = "Cache";
			// 
			// cacheGrid
			// 
			this.cacheGrid.AllowNavigation = false;
			this.cacheGrid.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.cacheGrid.CaptionText = "Cache";
			this.cacheGrid.DataMember = "";
			this.cacheGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.cacheGrid.Location = new System.Drawing.Point(8, 8);
			this.cacheGrid.Name = "cacheGrid";
			this.cacheGrid.ReadOnly = true;
			this.cacheGrid.Size = new System.Drawing.Size(682, 424);
			this.cacheGrid.TabIndex = 0;
			// 
			// tabSession
			// 
			this.tabSession.Controls.AddRange(new System.Windows.Forms.Control[] {
																					 this.sessionGrid});
			this.tabSession.Location = new System.Drawing.Point(4, 4);
			this.tabSession.Name = "tabSession";
			this.tabSession.Size = new System.Drawing.Size(698, 438);
			this.tabSession.TabIndex = 8;
			this.tabSession.Text = "Session";
			// 
			// sessionGrid
			// 
			this.sessionGrid.AllowNavigation = false;
			this.sessionGrid.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.sessionGrid.CaptionText = "Session";
			this.sessionGrid.DataMember = "";
			this.sessionGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.sessionGrid.Location = new System.Drawing.Point(8, 7);
			this.sessionGrid.Name = "sessionGrid";
			this.sessionGrid.ReadOnly = true;
			this.sessionGrid.Size = new System.Drawing.Size(682, 424);
			this.sessionGrid.TabIndex = 1;
			// 
			// tabViewState
			// 
			this.tabViewState.Controls.AddRange(new System.Windows.Forms.Control[] {
																					   this.viewstateGrid});
			this.tabViewState.Location = new System.Drawing.Point(4, 4);
			this.tabViewState.Name = "tabViewState";
			this.tabViewState.Size = new System.Drawing.Size(698, 438);
			this.tabViewState.TabIndex = 9;
			this.tabViewState.Text = "View State";
			// 
			// viewstateGrid
			// 
			this.viewstateGrid.AllowNavigation = false;
			this.viewstateGrid.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.viewstateGrid.CaptionText = "View State";
			this.viewstateGrid.DataMember = "";
			this.viewstateGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.viewstateGrid.Location = new System.Drawing.Point(8, 7);
			this.viewstateGrid.Name = "viewstateGrid";
			this.viewstateGrid.ReadOnly = true;
			this.viewstateGrid.Size = new System.Drawing.Size(682, 424);
			this.viewstateGrid.TabIndex = 2;
			// 
			// tabHeaders
			// 
			this.tabHeaders.Controls.AddRange(new System.Windows.Forms.Control[] {
																					 this.headersGrid});
			this.tabHeaders.Location = new System.Drawing.Point(4, 4);
			this.tabHeaders.Name = "tabHeaders";
			this.tabHeaders.Size = new System.Drawing.Size(698, 438);
			this.tabHeaders.TabIndex = 3;
			this.tabHeaders.Text = "Headers";
			// 
			// headersGrid
			// 
			this.headersGrid.AllowNavigation = false;
			this.headersGrid.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.headersGrid.CaptionText = "Request.Headers";
			this.headersGrid.DataMember = "";
			this.headersGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.headersGrid.Location = new System.Drawing.Point(8, 7);
			this.headersGrid.Name = "headersGrid";
			this.headersGrid.ReadOnly = true;
			this.headersGrid.Size = new System.Drawing.Size(682, 424);
			this.headersGrid.TabIndex = 3;
			// 
			// tabRequest
			// 
			this.tabRequest.Controls.AddRange(new System.Windows.Forms.Control[] {
																					 this.formGrid});
			this.tabRequest.Location = new System.Drawing.Point(4, 4);
			this.tabRequest.Name = "tabRequest";
			this.tabRequest.Size = new System.Drawing.Size(698, 438);
			this.tabRequest.TabIndex = 4;
			this.tabRequest.Text = "Request";
			// 
			// formGrid
			// 
			this.formGrid.AllowNavigation = false;
			this.formGrid.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.formGrid.CaptionText = "Request.Form";
			this.formGrid.DataMember = "";
			this.formGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.formGrid.Location = new System.Drawing.Point(8, 7);
			this.formGrid.Name = "formGrid";
			this.formGrid.ReadOnly = true;
			this.formGrid.Size = new System.Drawing.Size(682, 424);
			this.formGrid.TabIndex = 1;
			// 
			// tabCookies
			// 
			this.tabCookies.Controls.AddRange(new System.Windows.Forms.Control[] {
																					 this.cookiesGrid});
			this.tabCookies.Location = new System.Drawing.Point(4, 4);
			this.tabCookies.Name = "tabCookies";
			this.tabCookies.Size = new System.Drawing.Size(698, 438);
			this.tabCookies.TabIndex = 6;
			this.tabCookies.Text = "Cookies";
			// 
			// cookiesGrid
			// 
			this.cookiesGrid.AllowNavigation = false;
			this.cookiesGrid.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.cookiesGrid.CaptionText = "Request.Cookies";
			this.cookiesGrid.DataMember = "";
			this.cookiesGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.cookiesGrid.Location = new System.Drawing.Point(8, 7);
			this.cookiesGrid.Name = "cookiesGrid";
			this.cookiesGrid.ReadOnly = true;
			this.cookiesGrid.Size = new System.Drawing.Size(682, 424);
			this.cookiesGrid.TabIndex = 2;
			// 
			// tabServerVars
			// 
			this.tabServerVars.Controls.AddRange(new System.Windows.Forms.Control[] {
																						this.serverGrid});
			this.tabServerVars.Location = new System.Drawing.Point(4, 4);
			this.tabServerVars.Name = "tabServerVars";
			this.tabServerVars.Size = new System.Drawing.Size(698, 438);
			this.tabServerVars.TabIndex = 5;
			this.tabServerVars.Text = "Server Vars";
			// 
			// serverGrid
			// 
			this.serverGrid.AllowNavigation = false;
			this.serverGrid.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.serverGrid.CaptionText = "Request.ServerVariables";
			this.serverGrid.DataMember = "";
			this.serverGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.serverGrid.Location = new System.Drawing.Point(8, 7);
			this.serverGrid.Name = "serverGrid";
			this.serverGrid.ReadOnly = true;
			this.serverGrid.Size = new System.Drawing.Size(682, 424);
			this.serverGrid.TabIndex = 2;
			// 
			// tabPageControls
			// 
			this.tabPageControls.Controls.AddRange(new System.Windows.Forms.Control[] {
																						  this.formControlsGrid,
																						  this.controlsGrid});
			this.tabPageControls.Location = new System.Drawing.Point(4, 4);
			this.tabPageControls.Name = "tabPageControls";
			this.tabPageControls.Size = new System.Drawing.Size(698, 438);
			this.tabPageControls.TabIndex = 7;
			this.tabPageControls.Text = "Page Controls";
			// 
			// formControlsGrid
			// 
			this.formControlsGrid.AllowNavigation = false;
			this.formControlsGrid.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.formControlsGrid.CaptionText = "Form.Controls";
			this.formControlsGrid.DataMember = "";
			this.formControlsGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.formControlsGrid.Location = new System.Drawing.Point(8, 184);
			this.formControlsGrid.Name = "formControlsGrid";
			this.formControlsGrid.ReadOnly = true;
			this.formControlsGrid.Size = new System.Drawing.Size(682, 248);
			this.formControlsGrid.TabIndex = 4;
			// 
			// controlsGrid
			// 
			this.controlsGrid.AllowNavigation = false;
			this.controlsGrid.Anchor = ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.controlsGrid.CaptionText = "Page.Controls";
			this.controlsGrid.DataMember = "";
			this.controlsGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.controlsGrid.Location = new System.Drawing.Point(8, 7);
			this.controlsGrid.Name = "controlsGrid";
			this.controlsGrid.ReadOnly = true;
			this.controlsGrid.Size = new System.Drawing.Size(682, 177);
			this.controlsGrid.TabIndex = 3;
			// 
			// appStatusBar
			// 
			this.appStatusBar.Location = new System.Drawing.Point(0, 511);
			this.appStatusBar.Name = "appStatusBar";
			this.appStatusBar.Size = new System.Drawing.Size(722, 22);
			this.appStatusBar.TabIndex = 1;
			this.appStatusBar.Text = "Ready";
			// 
			// buttonGo
			// 
			this.buttonGo.Anchor = (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.buttonGo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.buttonGo.Location = new System.Drawing.Point(682, 8);
			this.buttonGo.Name = "buttonGo";
			this.buttonGo.Size = new System.Drawing.Size(32, 20);
			this.buttonGo.TabIndex = 5;
			this.buttonGo.Text = "Go";
			this.buttonGo.Click += new System.EventHandler(this.buttonGo_Click);
			// 
			// addressBar
			// 
			this.addressBar.Anchor = ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.addressBar.BackColor = System.Drawing.SystemColors.Info;
			this.addressBar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.addressBar.Location = new System.Drawing.Point(8, 8);
			this.addressBar.Name = "addressBar";
			this.addressBar.Size = new System.Drawing.Size(674, 20);
			this.addressBar.TabIndex = 4;
			this.addressBar.Text = "http://localhost/mydebug/products.aspx";
			// 
			// MyTracerForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(722, 533);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.buttonGo,
																		  this.addressBar,
																		  this.appStatusBar,
																		  this.tabs});
			this.MinimumSize = new System.Drawing.Size(730, 560);
			this.Name = "MyTracerForm";
			this.Text = "My Tracer";
			this.tabs.ResumeLayout(false);
			this.tabBrowser.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.webBrowser)).EndInit();
			this.tabApplication.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.appGrid)).EndInit();
			this.tabCache.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cacheGrid)).EndInit();
			this.tabSession.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.sessionGrid)).EndInit();
			this.tabViewState.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.viewstateGrid)).EndInit();
			this.tabHeaders.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.headersGrid)).EndInit();
			this.tabRequest.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.formGrid)).EndInit();
			this.tabCookies.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cookiesGrid)).EndInit();
			this.tabServerVars.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.serverGrid)).EndInit();
			this.tabPageControls.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.formControlsGrid)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.controlsGrid)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion



		public MyTracerForm()
		{
			InitApp();
		}
		
		public MyTracerForm(string url) : base()
		{
			InitApp();
			addressBar.Text = url;
			buttonGo_Click(this, EventArgs.Empty); 
		}

		private void InitApp()
		{
			InitializeComponent();
			
			// Navigate to a default help text
			NavigateUrl(null);
			UpdateEnvInfo(null, true);

			// Hook up the StatusTextChange event
			DWebBrowserEvents2_StatusTextChangeEventHandler eh_StatusTextChange = new DWebBrowserEvents2_StatusTextChangeEventHandler(StatusTextChange);
			webBrowser.StatusTextChange += eh_StatusTextChange;

			// Hook up the DocumentComplete event
			DWebBrowserEvents2_DocumentCompleteEventHandler eh_DocumentComplete = new DWebBrowserEvents2_DocumentCompleteEventHandler(DocumentComplete);
			webBrowser.DocumentComplete += eh_DocumentComplete;

		}

		[STAThread]
		static void Main(string[] args) 
		{
			if (args.Length >0)
				Application.Run(new MyTracerForm(args[0]));
			else
				Application.Run(new MyTracerForm());
		}


		// *********************************************************************
		// Make the WebBrowser to navigate to the specified URL
		private void NavigateUrl(string url)
		{
			string address = APP_DEFAULTURL;
			if (url != null)
			{
				address = url;
				//addressBar.Text = url;
			}
						
			object o1 = null, o2 = null, o3 = null, o4 = null;
			webBrowser.Navigate(address, ref o1, ref o2, ref o3, ref o4);
		}
		// *********************************************************************

		// *********************************************************************
		// Run when the Go button is clicked
		private void buttonGo_Click(object sender, System.EventArgs e)
		{
			UpdateBrowserTab();
		}
		// *********************************************************************

		// *********************************************************************
		// Navigate to a new page and update the view state tab
		private void UpdateBrowserTab()
		{
			NavigateUrl(addressBar.Text);
		}
		// *********************************************************************

		// *********************************************************************
		// Handle the WebBrowser's status text change event
		private void StatusTextChange(object sender, DWebBrowserEvents2_StatusTextChangeEvent e)
		{
			appStatusBar.Text = e.text;
		}
		// *********************************************************************

		// *********************************************************************
		// Handle the WebBrowser's document complete event
		private void DocumentComplete(object sender, DWebBrowserEvents2_DocumentCompleteEvent e)
		{
			if (APP_DEFAULTURL == e.uRL.ToString())
				return;

			UpdateEnvInfo(e.uRL.ToString(), false);
			UpdatePageInfo();
		}
		// *********************************************************************

		// *********************************************************************
		// Refresh the UI of the program
		private void UpdateEnvInfo(string url, bool resetUI)
		{
			m_currentPage = url;

			if (!resetUI)
				this.Text = String.Format(APP_TITLE + " [{0}]", m_currentPage);
			else
				this.Text = APP_TITLE; 
		}
		// *********************************************************************


		// *********************************************************************
		// Decode and display the view state information
		private void UpdatePageInfo()
		{
			MyTracer.MyDebug.MyDebugTool tool = new MyTracer.MyDebug.MyDebugTool();
			DataSet ds = tool.GetInfo(CONNSTRING, "dino");	

			cacheGrid.DataSource = ds.Tables["Cache"];
			appGrid.DataSource = ds.Tables["Application"];
			headersGrid.DataSource = ds.Tables["Headers"];
			serverGrid.DataSource = ds.Tables["ServerVariables"];
			cookiesGrid.DataSource = ds.Tables["Cookies"];
			formGrid.DataSource = ds.Tables["Form"];
			sessionGrid.DataSource = ds.Tables["Session"];
			viewstateGrid.DataSource = ds.Tables["ViewState"];
			controlsGrid.DataSource = ds.Tables["Controls"];
			formControlsGrid.DataSource = ds.Tables["FormControls"];
		}
		// *********************************************************************
	}
}
