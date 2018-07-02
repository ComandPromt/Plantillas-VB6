// MfcWinFormsHostDlg.cpp : implementation file
//

#include "stdafx.h"
#include "MfcWinFormsHost.h"
#include "MfcWinFormsHostDlg.h"
#include ".\mfcwinformshostdlg.h"

// DONE: Can't use this and create instances of managed types
//#ifdef _DEBUG
//#define new DEBUG_NEW
//#endif

// DONE: Bring in the System.Windows.Forms assembly
#using <System.dll>
#using <System.Windows.Forms.dll>

// CMfcWinFormsHostDlg dialog



CMfcWinFormsHostDlg::CMfcWinFormsHostDlg(CWnd* pParent /*=NULL*/)
: CDialog(CMfcWinFormsHostDlg::IDD, pParent)
{
  m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMfcWinFormsHostDlg::DoDataExchange(CDataExchange* pDX)
{
  CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CMfcWinFormsHostDlg, CDialog)
  ON_WM_PAINT()
  ON_WM_QUERYDRAGICON()
  //}}AFX_MSG_MAP
  ON_WM_CREATE()
END_MESSAGE_MAP()


// CMfcWinFormsHostDlg message handlers

BOOL CMfcWinFormsHostDlg::OnInitDialog()
{
  CDialog::OnInitDialog();

  // Set the icon for this dialog.  The framework does this automatically
  //  when the application's main window is not a dialog
  SetIcon(m_hIcon, TRUE);			// Set big icon
  SetIcon(m_hIcon, FALSE);		// Set small icon

  // DONE: Create WinForms control(s)
  System::Windows::Forms::MonthCalendar* pcal = new System::Windows::Forms::MonthCalendar();

  // DONE: Get interface pointer from control
  CComPtr<IUnknown> spunkControl;
  spunkControl.Attach((IUnknown*)System::Runtime::InteropServices::Marshal::GetIUnknownForObject(pcal).ToPointer());

  // DONE: Get rect of placeholder control and then hide it
  CWnd* pwnd = this->GetDlgItem(IDC_WINFORMS_MONTH_CALENDAR_PLACEHOLDER);
  CRect rectPlaceHolder;
  pwnd->GetWindowRect(&rectPlaceHolder);
  this->ScreenToClient(rectPlaceHolder);
  pwnd->ShowWindow(SW_HIDE);

  // DONE: Wrap control and place in on the parent window
  m_wndWinFormsCalendar.Create(spunkControl, WS_VISIBLE | WS_TABSTOP, rectPlaceHolder, this, 0);

  return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CMfcWinFormsHostDlg::OnPaint() 
{
  if (IsIconic())
  {
    CPaintDC dc(this); // device context for painting

    SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

    // Center icon in client rectangle
    int cxIcon = GetSystemMetrics(SM_CXICON);
    int cyIcon = GetSystemMetrics(SM_CYICON);
    CRect rect;
    GetClientRect(&rect);
    int x = (rect.Width() - cxIcon + 1) / 2;
    int y = (rect.Height() - cyIcon + 1) / 2;

    // Draw the icon
    dc.DrawIcon(x, y, m_hIcon);
  }
  else
  {
    CDialog::OnPaint();
  }
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CMfcWinFormsHostDlg::OnQueryDragIcon()
{
  return static_cast<HCURSOR>(m_hIcon);
}

// DONE: Hand our our custom control site
BOOL CMfcWinFormsHostDlg::CreateControlSite(COleControlContainer* pContainer, COleControlSite** ppSite, UINT nID, REFCLSID clsid)
{
  ASSERT(ppSite);
  *ppSite = new CWinFormsControlSite(this->GetControlContainer());
  return TRUE;
}












