////////////////////////////////////////////////////////////////
// MSDN Magazine -- March 2003
// If this code works, it was written by Paul DiLascia.
// If not, I don't know who wrote it.
// Compiles with Visual Studio .NET on Windows XP. Tab size=3.
//
// See SysIcons.CPP for description of program.
// 
#include "resource.h"

class CMyApp : public CWinApp {
public:
	CMyApp();
	virtual BOOL InitInstance();
	afx_msg void OnAppAbout();
	DECLARE_MESSAGE_MAP()
};
