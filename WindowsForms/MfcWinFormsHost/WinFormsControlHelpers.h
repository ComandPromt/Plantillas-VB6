// ManagedControlHelpers.h: A set of classes for hosting WinForms controls in MFC7.x.
// Based on work by Mark Boulter.
// Copyright (c) 2002, Chris Sells.
// Feel free to use as you like.
// No warranties extended. Use at your own risk.

#ifndef __MANAGEDCONTROLHELPERS_H__
#define __MANAGEDCONTROLHELPERS_H__

#pragma once
#include <afxocc.h>
#include <atlcomcli.h>
#ifndef S_QUICKACTIVATED
#define S_QUICKACTIVATED S_FALSE
#endif

// Keep the HWND and the IUnknown for the managed control together
class CWinFormsControlWnd : public CWnd {
public:
  CWinFormsControlWnd()
  {
    // Initialize OLE, if necessary
    _AFX_THREAD_STATE* pState = AfxGetThreadState();
    if( !pState->m_bNeedTerm && !AfxOleInit() ) ASSERT(FALSE);
  }

  BOOL Create(
    IUnknown* punkControl,
    DWORD dwStyle,
    const RECT& rect,
    CWnd* pParentWnd,
    UINT nID) {
      // Cache control interface pointer
      spunkControl = punkControl;

      // This call will end up in CWinFormsControlSite::CreateControl,
      // which is why it's OK to pass a NULL CLSID (we're wrapping a
      // pre-existing WinForms control, not creating a new COM control)
      return CreateControl(CLSID_NULL, 0, dwStyle, rect, pParentWnd, nID);
    }

    void SetControlSite(COleControlSite* pCtrlSite) {
      ASSERT(this->m_pCtrlSite == NULL);
      this->m_pCtrlSite = pCtrlSite;
    }

    HRESULT GetControl(IUnknown** ppunk) {
      if( !ppunk ) return E_POINTER;
      if( !spunkControl ) return E_UNEXPECTED;
      return spunkControl->QueryInterface(ppunk);
    }

private:
  CComPtr<IUnknown> spunkControl;
};

// A control site that knows how to wrap existing WinForms controls
class CWinFormsControlSite : public COleControlSite {
public:
  CWinFormsControlSite(COleControlContainer* pCtrlCont) : COleControlSite(pCtrlCont) {}

  // NOTE: This implementation is copied from occsite.cpp except for two changes
  virtual HRESULT CreateControl(
    CWnd* pWndCtrl,
    REFCLSID clsid,
    LPCTSTR lpszWindowName,
    DWORD dwStyle,
    const POINT* ppt,
    const SIZE* psize,
    UINT nID,
    CFile* pPersist = 0,
    BOOL bStorage = FALSE,
    BSTR bstrLicKey = 0) {
      HRESULT hr = E_FAIL;
      m_hWnd = NULL;
      CSize size;

      // Connect the OLE Control with its proxy CWnd object
      if (pWndCtrl != NULL)
      {
        // 1st change: Set the control's site using a member function,
        // not a protected CWnd variable
        CWinFormsControlWnd* pwnd = (CWinFormsControlWnd*)pWndCtrl;
        ASSERT(pwnd);
        pwnd->SetControlSite(this);
        //ASSERT(pWndCtrl->m_pCtrlSite == NULL);
        m_pWndCtrl = pWndCtrl;
        //pWndCtrl->m_pCtrlSite = this;
      }

      // Initialize OLE, if necessary
      _AFX_THREAD_STATE* pState = AfxGetThreadState();
      if (!pState->m_bNeedTerm && !AfxOleInit())
        return hr;

      // 2nd change: Wrap existing WinForms control
      // instead of creating a new COM control
      if (SUCCEEDED(hr = WrapWinFormsControl(clsid, pPersist, bStorage, bstrLicKey)))
      {
        ASSERT(m_pObject != NULL);
        m_nID = nID;

        if (psize == NULL)
        {
          // If psize is NULL, ask the object how big it wants to be.
          CClientDC dc(NULL);

          m_pObject->GetExtent(DVASPECT_CONTENT, &size);
          dc.HIMETRICtoDP(&size);
          m_rect = CRect(*ppt, size);
        }
        else
        {
          m_rect = CRect(*ppt, *psize);
        }

        m_dwStyleMask = WS_GROUP | WS_TABSTOP;

        if (m_dwMiscStatus & OLEMISC_ACTSLIKEBUTTON)
          m_dwStyleMask |= BS_DEFPUSHBUTTON;

        if (m_dwMiscStatus & OLEMISC_INVISIBLEATRUNTIME)
          dwStyle &= ~WS_VISIBLE;

        m_dwStyle = dwStyle & m_dwStyleMask;

        // If control wasn't quick-activated, then connect sinks now.
        if (hr != S_QUICKACTIVATED)
        {
          m_dwEventSink = ConnectSink(m_iidEvents, &m_xEventSink);
          m_dwPropNotifySink = ConnectSink(IID_IPropertyNotifySink, &m_xPropertyNotifySink);
        }
        m_dwNotifyDBEvents = ConnectSink(IID_INotifyDBEvents, &m_xNotifyDBEvents);

        // Now that the object has been created, attempt to
        // in-place activate it.

        SetExtent();

        if (SUCCEEDED(hr = m_pObject->QueryInterface(IID_IOleInPlaceObject, (LPVOID*)&m_pInPlaceObject)))
        {
          if (dwStyle & WS_VISIBLE)
          {
            // control is visible: just activate it
            hr = DoVerb(OLEIVERB_INPLACEACTIVATE);
          }
          else
          {
            // control is not visible: activate off-screen, hide, then move
            m_rect.OffsetRect(-32000, -32000);
            if (SUCCEEDED(hr = DoVerb(OLEIVERB_INPLACEACTIVATE)) &&
              SUCCEEDED(hr = DoVerb(OLEIVERB_HIDE)))
            {
              m_rect.OffsetRect(32000, 32000);
              hr = m_pInPlaceObject->SetObjectRects(m_rect, m_rect);
            }
          }
        }
        else
        {
          TRACE(traceOle, 0, "IOleInPlaceObject not supported on OLE control (dialog ID %d).\n", nID);
          TRACE(traceOle, 0, ">>> Result code: 0x%08lx\n", hr);
        }

        if (SUCCEEDED(hr))
          GetControlInfo();

        // if QueryInterface or activation failed, cleanup everything
        if (FAILED(hr))
        {
          if (m_pInPlaceObject != NULL)
          {
            m_pInPlaceObject->Release();
            m_pInPlaceObject = NULL;
          }
          DisconnectSink(m_iidEvents, m_dwEventSink);
          DisconnectSink(IID_IPropertyNotifySink, m_dwPropNotifySink);
          DisconnectSink(IID_INotifyDBEvents, m_dwNotifyDBEvents);
          m_dwEventSink = 0;
          m_dwPropNotifySink = 0;
          m_dwNotifyDBEvents = 0;
          m_pObject->Release();
          m_pObject = NULL;
        }
      }

      if (SUCCEEDED(hr))
      {
        AttachWindow();

        //		ASSERT(m_hWnd != NULL);

        // Initialize the control's Caption or Text property, if any
        if (lpszWindowName != NULL) SetWindowText(lpszWindowName);

        // Initialize styles
        ModifyStyle(0, m_dwStyle | (dwStyle & (WS_DISABLED|WS_BORDER)), 0);
      }

      return hr;
    }

protected:
  // NOTE: This code is based on COleControlSite::CreateOrLoad,
  // but without support for persistence or licensing
  virtual HRESULT WrapWinFormsControl(
    REFCLSID clsid,
    CFile* pPersist,
    BOOL bStorage,
    BSTR bstrLicKey) {
      HRESULT hr = E_FAIL;
      ASSERT(m_pObject == NULL);
      ASSERT(clsid == CLSID_NULL);
      ASSERT(!pPersist && "No support for persistence");
      ASSERT(!bStorage && "No support for structured storage");
      ASSERT(!bstrLicKey && "No support for licensing");

      CWinFormsControlWnd* pWndCtrlLocal = (CWinFormsControlWnd*)m_pWndCtrl;

      CComPtr<IUnknown> spunk;
      hr = pWndCtrlLocal->GetControl(&spunk);

      if (FAILED(hr = spunk->QueryInterface(IID_IOleObject, (void**)&m_pObject))) {
        return hr;
      }

      GetEventIID(&m_iidEvents);

      // Try to quick-activate first
      BOOL bQuickActivated = QuickActivate();

      if (!bQuickActivated) {
        m_pObject->GetMiscStatus(DVASPECT_CONTENT, &m_dwMiscStatus);

        // set client site first, if appropriate
        if (m_dwMiscStatus & OLEMISC_SETCLIENTSITEFIRST) {
          if (FAILED(hr = m_pObject->SetClientSite(&m_xOleClientSite))) {
            goto WrapFailed;
          }
        }
      }

      if (!bQuickActivated) {
        // set client site last, if appropriate
        if (!(m_dwMiscStatus & OLEMISC_SETCLIENTSITEFIRST)) {
          if (FAILED(hr = m_pObject->SetClientSite(&m_xOleClientSite))) {
            goto WrapFailed;
          }
        }
      }

WrapFailed:
      if (FAILED(hr) && (m_pObject != NULL)) {
        m_pObject->Close(OLECLOSE_NOSAVE);
        m_pObject->Release();
        m_pObject = NULL;
      }

      if (bQuickActivated && SUCCEEDED(hr)) hr = S_QUICKACTIVATED;
      return hr;
    }
};

#endif // __MANAGEDCONTROLHELPERS_H__











