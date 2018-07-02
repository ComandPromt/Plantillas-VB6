// anotherMfcTestView.cpp : implementation of the CanotherMfcTestView class
//

#include "stdafx.h"
#include "anotherMfcTest.h"

#include "anotherMfcTestDoc.h"
#include "anotherMfcTestView.h"
#include ".\anothermfctestview.h"

// DONE: Can't use this and create instances of managed types
//#ifdef _DEBUG
//#define new DEBUG_NEW
//#endif

// DONE: Bring in the System.Windows.Forms assembly
#using <System.dll>
#using <System.Windows.Forms.dll>

// CanotherMfcTestView

IMPLEMENT_DYNCREATE(CanotherMfcTestView, CView)

BEGIN_MESSAGE_MAP(CanotherMfcTestView, CView)
  // Standard printing commands
  ON_COMMAND(ID_FILE_PRINT, CView::OnFilePrint)
  ON_COMMAND(ID_FILE_PRINT_DIRECT, CView::OnFilePrint)
  ON_COMMAND(ID_FILE_PRINT_PREVIEW, CView::OnFilePrintPreview)
  ON_WM_SIZE()
  ON_WM_CREATE()
END_MESSAGE_MAP()

// CanotherMfcTestView construction/destruction

CanotherMfcTestView::CanotherMfcTestView()
{
  // TODO: add construction code here

}

CanotherMfcTestView::~CanotherMfcTestView()
{
}

BOOL CanotherMfcTestView::PreCreateWindow(CREATESTRUCT& cs)
{
  // TODO: Modify the Window class or styles here by modifying
  //  the CREATESTRUCT cs

  return CView::PreCreateWindow(cs);
}

// CanotherMfcTestView drawing

void CanotherMfcTestView::OnDraw(CDC* /*pDC*/)
{
  CanotherMfcTestDoc* pDoc = GetDocument();
  ASSERT_VALID(pDoc);
  if (!pDoc)
    return;

  // TODO: add draw code for native data here
}


// CanotherMfcTestView printing

BOOL CanotherMfcTestView::OnPreparePrinting(CPrintInfo* pInfo)
{
  // default preparation
  return DoPreparePrinting(pInfo);
}

void CanotherMfcTestView::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
  // TODO: add extra initialization before printing
}

void CanotherMfcTestView::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
  // TODO: add cleanup after printing
}


// CanotherMfcTestView diagnostics

#ifdef _DEBUG
void CanotherMfcTestView::AssertValid() const
{
  CView::AssertValid();
}

void CanotherMfcTestView::Dump(CDumpContext& dc) const
{
  CView::Dump(dc);
}

CanotherMfcTestDoc* CanotherMfcTestView::GetDocument() const // non-debug version is inline
{
  ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CanotherMfcTestDoc)));
  return (CanotherMfcTestDoc*)m_pDocument;
}
#endif //_DEBUG


// CanotherMfcTestView message handlers


// DONE: Hand our our custom control site
BOOL CanotherMfcTestView::CreateControlSite(COleControlContainer* pContainer, COleControlSite** ppSite, UINT nID, REFCLSID clsid)
{
  ASSERT(ppSite);
  *ppSite = new CWinFormsControlSite(this->GetControlContainer());
  return TRUE;
}

int CanotherMfcTestView::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
  if (CView::OnCreate(lpCreateStruct) == -1)
    return -1;

  // DONE: Create WinForms control(s)
  System::Windows::Forms::MonthCalendar* pcal = new System::Windows::Forms::MonthCalendar();

  // DONE: Get interface pointer from control
  CComPtr<IUnknown> spunkControl;
  spunkControl.Attach((IUnknown*)System::Runtime::InteropServices::Marshal::GetIUnknownForObject(pcal).ToPointer());

  // DONE: Get rect of client area
  CRect rect; this->GetClientRect(&rect);

  // DONE: Wrap control and place in on the parent window
  m_wndWinFormsCalendar.Create(spunkControl, WS_VISIBLE | WS_TABSTOP, rect, this, 0);

  return 0;
}

void CanotherMfcTestView::OnSize(UINT nType, int cx, int cy)
{
  CView::OnSize(nType, cx, cy);

  // DONE: Resize calendar to fit client
  if( !m_wndWinFormsCalendar.GetSafeHwnd() ) return;
  m_wndWinFormsCalendar.MoveWindow(0, 0, cx, cy);
}

