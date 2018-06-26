////////////////////////////////////////////////////////////////
// TRACEWIN Copyright 1997-1998 Paul DiLascia
// If this code works, it was written by Paul DiLascia.
// If not, I don't know who wrote it.
//
// Implementation of TRACEWIN applet. This is the applet sits around
// and waiting for WM_COPYDATA messages from other apps using TRACEWIN.
//
#include "StdAfx.h"
#include "MainFrm.h"
#include "resource.h"
#include "WinPlace.h"
#include "StatLink.h"

/////////////////
// Application class
//
class CTraceWinApp : public CWinApp {
public:
	CTraceWinApp();
	~CTraceWinApp();
	virtual BOOL InitInstance();
	afx_msg void OnAppAbout();
	DECLARE_MESSAGE_MAP()
};

CTraceWinApp NEAR theApp;

BEGIN_MESSAGE_MAP(CTraceWinApp, CWinApp)
	//{{AFX_MSG_MAP(CTraceWinApp)
	ON_COMMAND(ID_APP_ABOUT, OnAppAbout)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

CTraceWinApp::CTraceWinApp()
{
}

CTraceWinApp::~CTraceWinApp()
{
}

//////////////////
// TRACEWIN applet starting up.
//
BOOL CTraceWinApp::InitInstance()
{
	CWnd* pWnd = CWnd::FindWindow(_T(TRACEWND_CLASSNAME), NULL);
	if (pWnd) {
		pWnd->SetForegroundWindow();
		return FALSE;
	}

	// Save settings in registry, not INI file
	SetRegistryKey(_T("PixieLib"));

   // Create main frame window (don't use doc/view)
   CMainFrame* pMainFrame = new CMainFrame;
   if (!pMainFrame->LoadFrame(IDR_MAINFRAME))
      return FALSE;

	// Load window position from profile
	CWindowPlacement wp;
	if (!wp.Restore(CMainFrame::REGKEY, _T("MainWindow"), pMainFrame))
		pMainFrame->ShowWindow(m_nCmdShow);
   pMainFrame->UpdateWindow();
   m_pMainWnd = pMainFrame;

	return TRUE;
}

//////////////////
// Custom about dialog uses CStaticLink for hyperlinks.
//    * for text control, URL is specified as text in dialog editor
//    * for icon control, URL is specified by setting m_iconLink.m_link
//
class CAboutDialog : public CDialog {
protected:
	// static controls with hyperlinks
	CStaticLink	m_wndLink1;
	CStaticLink	m_wndLink2;
public:
	CAboutDialog() : CDialog(IDD_ABOUTBOX) { }
	virtual BOOL OnInitDialog()
	{
		// subclass static controls. URL is static text or 3rd arg
		m_wndLink1.SubclassDlgItem(IDC_URLTEXT,this);
		m_wndLink2.SubclassDlgItem(IDC_URLICON,this);
		return CDialog::OnInitDialog();
	}
};

//////////////////
// Handle Help | About : run the About dialog
//
void CTraceWinApp::OnAppAbout()
{
	static CAboutDialog dlg;
	dlg.DoModal();
}
