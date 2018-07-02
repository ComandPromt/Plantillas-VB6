// anotherMfcTestDoc.cpp : implementation of the CanotherMfcTestDoc class
//

#include "stdafx.h"
#include "anotherMfcTest.h"

#include "anotherMfcTestDoc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CanotherMfcTestDoc

IMPLEMENT_DYNCREATE(CanotherMfcTestDoc, CDocument)

BEGIN_MESSAGE_MAP(CanotherMfcTestDoc, CDocument)
END_MESSAGE_MAP()


// CanotherMfcTestDoc construction/destruction

CanotherMfcTestDoc::CanotherMfcTestDoc()
{
	// TODO: add one-time construction code here

}

CanotherMfcTestDoc::~CanotherMfcTestDoc()
{
}

BOOL CanotherMfcTestDoc::OnNewDocument()
{
	if (!CDocument::OnNewDocument())
		return FALSE;

	// TODO: add reinitialization code here
	// (SDI documents will reuse this document)

	return TRUE;
}




// CanotherMfcTestDoc serialization

void CanotherMfcTestDoc::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{
		// TODO: add storing code here
	}
	else
	{
		// TODO: add loading code here
	}
}


// CanotherMfcTestDoc diagnostics

#ifdef _DEBUG
void CanotherMfcTestDoc::AssertValid() const
{
	CDocument::AssertValid();
}

void CanotherMfcTestDoc::Dump(CDumpContext& dc) const
{
	CDocument::Dump(dc);
}
#endif //_DEBUG


// CanotherMfcTestDoc commands
