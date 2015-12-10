// DlgProxy.h : header file
//

#if !defined(AFX_DLGPROXY_H__45CAA10F_15CE_4C35_B85E_7209B6F47CB1__INCLUDED_)
#define AFX_DLGPROXY_H__45CAA10F_15CE_4C35_B85E_7209B6F47CB1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CTestDlg;

/////////////////////////////////////////////////////////////////////////////
// CTestDlgAutoProxy command target

class CTestDlgAutoProxy : public CCmdTarget
{
	DECLARE_DYNCREATE(CTestDlgAutoProxy)

	CTestDlgAutoProxy();           // protected constructor used by dynamic creation

// Attributes
public:
	CTestDlg* m_pDialog;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTestDlgAutoProxy)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	virtual ~CTestDlgAutoProxy();

	// Generated message map functions
	//{{AFX_MSG(CTestDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(CTestDlgAutoProxy)

	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CTestDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGPROXY_H__45CAA10F_15CE_4C35_B85E_7209B6F47CB1__INCLUDED_)
