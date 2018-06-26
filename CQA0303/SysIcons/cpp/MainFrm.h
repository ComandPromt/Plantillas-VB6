////////////////////////////////////////////////////////////////
// MSDN Magazine -- March 2003
// If this code works, it was written by Paul DiLascia.
// If not, I don't know who wrote it.
// Compiles with Visual Studio .NET on Windows XP. Tab size=3.
//
class CMainFrame : public CFrameWnd {
public:
	CMainFrame();
	virtual ~CMainFrame();
protected:
	DECLARE_DYNAMIC(CMainFrame)
	CStatusBar  m_wndStatusBar;
	CToolBar    m_wndToolBar;
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnPaint();
	DECLARE_MESSAGE_MAP()
};
