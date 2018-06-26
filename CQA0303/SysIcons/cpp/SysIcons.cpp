////////////////////////////////////////////////////////////////
// MSDN Magazine -- March 2003
// If this code works, it was written by Paul DiLascia.
// If not, I don't know who wrote it.
// Compiles with Visual Studio .NET on Windows XP. Tab size=3.
//
// SysIcons illustrates xxxx
// Compiles with

#include "StdAfx.h"
#include "SysIcons.h"
#include "MainFrm.h"
#include "StatLink.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

CMyApp NEAR theApp;

BEGIN_MESSAGE_MAP(CMyApp, CWinApp)
	ON_COMMAND(ID_APP_ABOUT, OnAppAbout)
END_MESSAGE_MAP()

CMyApp::CMyApp()
{
}

BOOL CMyApp::InitInstance()
{
   // Create main frame window (don't use doc/view stuff)
   CMainFrame* pMainFrame = new CMainFrame;
   if (!pMainFrame->LoadFrame(IDR_MAINFRAME))
      return FALSE;
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
	CStaticLink	m_wndLink3;

public:
	CAboutDialog() : CDialog(IDD_ABOUTBOX) { }
	virtual BOOL OnInitDialog();
};

/////////////////
// Initialize dialog: subclass static text/icon controls
//
BOOL CAboutDialog::OnInitDialog()
{
	// subclass static controls. URL is static text or 3rd arg
	m_wndLink1.SubclassDlgItem(IDC_STATICURLPD,this);
	m_wndLink2.SubclassDlgItem(IDC_STATICURLMSDN,this);
	m_wndLink3.SubclassDlgItem(IDC_MSDNLINK,this);
	m_wndLink3.SetIcon(::LoadIcon(NULL, IDI_QUESTION));
	return CDialog::OnInitDialog();
}

//////////////////
// Handle Help | About : run the About dialog
//
void CMyApp::OnAppAbout()
{
	static CAboutDialog dlg;
	dlg.DoModal();
}
