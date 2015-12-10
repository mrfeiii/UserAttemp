// DlgProxy.cpp : implementation file
//

#include "stdafx.h"
#include "test.h"
#include "DlgProxy.h"
#include "testDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CTestDlgAutoProxy

IMPLEMENT_DYNCREATE(CTestDlgAutoProxy, CCmdTarget)

CTestDlgAutoProxy::CTestDlgAutoProxy()
{
	EnableAutomation();
	
	// To keep the application running as long as an automation 
	//	object is active, the constructor calls AfxOleLockApp.
	AfxOleLockApp();

	// Get access to the dialog through the application's
	//  main window pointer.  Set the proxy's internal pointer
	//  to point to the dialog, and set the dialog's back pointer to
	//  this proxy.
	ASSERT (AfxGetApp()->m_pMainWnd != NULL);
	ASSERT_VALID (AfxGetApp()->m_pMainWnd);
	ASSERT_KINDOF(CTestDlg, AfxGetApp()->m_pMainWnd);
	m_pDialog = (CTestDlg*) AfxGetApp()->m_pMainWnd;
	m_pDialog->m_pAutoProxy = this;
}

CTestDlgAutoProxy::~CTestDlgAutoProxy()
{
	// To terminate the application when all objects created with
	// 	with automation, the destructor calls AfxOleUnlockApp.
	//  Among other things, this will destroy the main dialog
	if (m_pDialog != NULL)
		m_pDialog->m_pAutoProxy = NULL;
	AfxOleUnlockApp();
}

void CTestDlgAutoProxy::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CTestDlgAutoProxy, CCmdTarget)
	//{{AFX_MSG_MAP(CTestDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CTestDlgAutoProxy, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CTestDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_ITest to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {8F0617DF-57C0-499C-B332-9A481F0FE9C5}
static const IID IID_ITest =
{ 0x8f0617df, 0x57c0, 0x499c, { 0xb3, 0x32, 0x9a, 0x48, 0x1f, 0xf, 0xe9, 0xc5 } };

BEGIN_INTERFACE_MAP(CTestDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CTestDlgAutoProxy, IID_ITest, Dispatch)
END_INTERFACE_MAP()

// The IMPLEMENT_OLECREATE2 macro is defined in StdAfx.h of this project
// {7F29BC88-1351-4E50-AAC6-D12AE21DE75A}
IMPLEMENT_OLECREATE2(CTestDlgAutoProxy, "Test.Application", 0x7f29bc88, 0x1351, 0x4e50, 0xaa, 0xc6, 0xd1, 0x2a, 0xe2, 0x1d, 0xe7, 0x5a)

/////////////////////////////////////////////////////////////////////////////
// CTestDlgAutoProxy message handlers
