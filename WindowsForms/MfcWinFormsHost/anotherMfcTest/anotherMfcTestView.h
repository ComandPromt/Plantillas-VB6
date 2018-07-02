// anotherMfcTestView.h : interface of the CanotherMfcTestView class
//


#pragma once
#include "..\WinFormsControlHelpers.h"

class CanotherMfcTestView : public CView
{
protected: // create from serialization only
	CanotherMfcTestView();
	DECLARE_DYNCREATE(CanotherMfcTestView)

// Attributes
public:
	CanotherMfcTestDoc* GetDocument() const;

// Operations
public:

// Overrides
	public:
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);

// Implementation
public:
	virtual ~CanotherMfcTestView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	DECLARE_MESSAGE_MAP()

  // DONE: Host WinForms controls as COM controls
  virtual BOOL CreateControlSite(COleControlContainer* pContainer, COleControlSite** ppSite, UINT nID, REFCLSID clsid);

private:
  // DONE: A wrapper for a WinForms control
  CWinFormsControlWnd m_wndWinFormsCalendar;

public:
  afx_msg void OnSize(UINT nType, int cx, int cy);
  afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
};

#ifndef _DEBUG  // debug version in anotherMfcTestView.cpp
inline CanotherMfcTestDoc* CanotherMfcTestView::GetDocument() const
   { return reinterpret_cast<CanotherMfcTestDoc*>(m_pDocument); }
#endif

