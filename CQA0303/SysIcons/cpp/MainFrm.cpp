////////////////////////////////////////////////////////////////
// MSDN Magazine -- March 2003
// If this code works, it was written by Paul DiLascia.
// If not, I don't know who wrote it.
// Compiles with Visual Studio .NET on Windows XP. Tab size=3.
//
#include "StdAfx.h"
#include "SysIcons.h"
#include "MainFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

IMPLEMENT_DYNAMIC(CMainFrame, CFrameWnd)

BEGIN_MESSAGE_MAP(CMainFrame, CFrameWnd)
	ON_WM_CREATE()
	ON_WM_PAINT()
END_MESSAGE_MAP()

CMainFrame::CMainFrame()
{
}

CMainFrame::~CMainFrame()
{
}

BOOL CMainFrame::PreCreateWindow(CREATESTRUCT& cs)
{
	BOOL bRet = CFrameWnd::PreCreateWindow(cs);
	cs.cx=200;
	cs.cy=400;
	return bRet;
}

int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	return CFrameWnd::OnCreate(lpCreateStruct);
}

const struct {
	LPCTSTR nResID;
	LPCTSTR name;
} SysIcons[] = {
	{ IDI_APPLICATION, _T("IDI_APPLICATION") },
	{ IDI_HAND, _T("IDI_HAND") },
	{ IDI_QUESTION, _T("IDI_QUESTION") },
	{ IDI_EXCLAMATION, _T("IDI_EXCLAMATION") },
	{ IDI_ASTERISK, _T("IDI_ASTERISK") },
#if(WINVER >= 0x0400)
	{ IDI_WINLOGO, _T("IDI_WINLOGO") },
	{ IDI_WARNING, _T("IDI_WARNING") },
	{ IDI_ERROR, _T("IDI_ERROR") },
	{ IDI_INFORMATION, _T("IDI_INFORMATION") },
#endif
	{ NULL, NULL }
};

void CMainFrame::OnPaint()
{
	CPaintDC dc(this);

	CRect rcClient;
	GetClientRect(&rcClient);

	int cyIcon = GetSystemMetrics(SM_CYICON);
	int cxIcon = GetSystemMetrics(SM_CXICON);

	CRect rcIcon(0,0,cxIcon,cyIcon);
	CRect rcText(cxIcon, 0, rcClient.Width()-cxIcon, cyIcon);

	for (UINT i=0; SysIcons[i].nResID; i++) {
		HICON hicon = ::LoadIcon(NULL, SysIcons[i].nResID);
		ASSERT(hicon);
		CString name = SysIcons[i].name;
		dc.DrawIcon(rcIcon.TopLeft(), hicon);
		dc.DrawText(name, rcText, DT_LEFT);
		rcIcon += CPoint(0, rcIcon.Height());
		rcText += CPoint(0, rcIcon.Height());
	}
}
