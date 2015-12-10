// testDlg.cpp : implementation file
//

#include "stdafx.h"
#include "test.h"
#include "testDlg.h"
#include "DlgProxy.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTestDlg dialog

IMPLEMENT_DYNAMIC(CTestDlg, CDialog);

CTestDlg::CTestDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CTestDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CTestDlg)
	m_StartNumber = _T("");
	m_CurrentNumber = _T("");
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;
}

CTestDlg::~CTestDlg()
{
	// If there is an automation proxy for this dialog, set
	//  its back pointer to this dialog to NULL, so it knows
	//  the dialog has been deleted.
	if (m_pAutoProxy != NULL)
		m_pAutoProxy->m_pDialog = NULL;
}

void CTestDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CTestDlg)
	DDX_Control(pDX, IDC_COMBO1, m_SelectArea);
	DDX_Control(pDX, IDC_BUTTON2, m_BtAutoFind);
	DDX_Control(pDX, IDC_EXPLORER1, m_ctrlWeb);
	DDX_Text(pDX, IDC_StartNumber, m_StartNumber);
	DDX_Text(pDX, IDC_CurentNumber, m_CurrentNumber);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CTestDlg, CDialog)
	//{{AFX_MSG_MAP(CTestDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_CLOSE()
	ON_BN_CLICKED(IDC_BUTTON1, OnWriteIn)
	ON_BN_CLICKED(IDC_BUTTON2, OnAutoFind)
	ON_WM_TIMER()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTestDlg message handlers

BOOL CTestDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

    m_SelectArea.AddString("甘肃");
    m_SelectArea.AddString("湖南");
	m_SelectArea.InsertString(0,"北京");
	m_SelectArea.AddString("青海");
	m_SelectArea.AddString("浙江");
	m_SelectArea.AddString("新疆");
    m_SelectArea.SetCurSel(0);

    m_ctrlWeb.Navigate("http://www.dlnu.edu.cn",NULL,NULL,NULL,NULL);

	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CTestDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CTestDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

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

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CTestDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

// Automation servers should not exit when a user closes the UI
//  if a controller still holds on to one of its objects.  These
//  message handlers make sure that if the proxy is still in use,
//  then the UI is hidden but the dialog remains around if it
//  is dismissed.

void CTestDlg::OnClose() 
{
	if (CanExit())
		CDialog::OnClose();
}

void CTestDlg::OnOK() 
{

	IHTMLDocument2 *pHTMLDocument=NULL;
	IHTMLElementCollection   *pAllElement=NULL;
    CString strJsFuc = "checkForm()";

	
	if (!(pHTMLDocument = (IHTMLDocument2*)m_ctrlWeb.GetDocument()))   //获取 IHTMLDocument2 的接口指针
		return;
    
	pHTMLDocument->get_all(&pAllElement); 
	CComPtr<IDispatch> pDispCommon;
	CComQIPtr<IHTMLInputTextElement,&IID_IHTMLInputTextElement>pElement;
	HRESULT hr=pAllElement->item(COleVariant("username"),COleVariant((long)0),&pDispCommon);
	
	if ((hr == S_OK) && (pDispCommon != NULL))
	{
        pElement=pDispCommon;
        CString str1="11111111111";
        pElement->put_value(str1.AllocSysString());
        pDispCommon.Release();
		
		hr=pAllElement->item(COleVariant("passwd"),COleVariant((long)0),&pDispCommon);
		
        pElement=pDispCommon;
        CString str2="123456";
        pElement->put_value(str2.AllocSysString());
        pDispCommon.Release();
		
		
		IHTMLWindow2*  pWindow;  
		pHTMLDocument->get_parentWindow(&pWindow);
		
		VARIANT  ret;  
		
		ret.vt    =  VT_EMPTY; 
		
		BSTR bstrCode = strJsFuc.AllocSysString();
		BSTR bstrLanguage = SysAllocString(L"javascript");
		pWindow->execScript(bstrCode, bstrLanguage, &ret);
       
		return;
	}
	else
	{
		ShellExecute(NULL, "open", "http://user.qzone.qq.com/316397938/main",NULL, NULL, SW_SHOWNORMAL); 
		return;
	}

	if (CanExit())
		CDialog::OnOK();
}

void CTestDlg::OnCancel() 
{
	if (CanExit())
		CDialog::OnCancel();
}

BOOL CTestDlg::CanExit()
{
	// If the proxy object is still around, then the automation
	//  controller is still holding on to this application.  Leave
	//  the dialog around, but hide its UI.
	if (m_pAutoProxy != NULL)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}

void CTestDlg::OnWriteIn() 
{
 UpdateData(TRUE);
	if(m_StartNumber.GetLength()==11)
	{	
		for(int i=0;i<11;i++)
		{
			if(m_StartNumber[i]>='0' && m_StartNumber[i]<='9')
			{}else
			{
				MessageBox("您输入的初始值不合法，必须为数字！");
				return;
			}
		}
		m_CurrentNumber	= m_StartNumber;
		UpdateData(FALSE);
	}else MessageBox("您输入的初始值有误，必须为11位！");	
}

void CTestDlg::OnAutoFind() 
{

CString name;
m_BtAutoFind.GetWindowText(name);
if(name=="自动查找")
{
	m_BtAutoFind.SetWindowText("停止查找");
	
   CButton *pBtn= (CButton *)GetDlgItem(IDC_BUTTON1); //IDC_BUTTON2这个按钮
   pBtn->EnableWindow(FALSE); // True or False
   pBtn= (CButton *)GetDlgItem(IDC_BUTTON3); //IDC_BUTTON2这个按钮
   pBtn->EnableWindow(FALSE);
   SetTimer(1,1000,NULL);

}else if(name=="停止查找")
{
    m_BtAutoFind.SetWindowText("自动查找");
	KillTimer(1);
   
	CButton *pBtn= (CButton *)GetDlgItem(IDC_BUTTON1);
    pBtn->EnableWindow(TRUE);
	pBtn= (CButton *)GetDlgItem(IDC_BUTTON3);
    pBtn->EnableWindow(TRUE);
} 


	
}

int CTestDlg::FillBlank(CString Number)
{
	IHTMLDocument2 *pHTMLDocument=NULL;
	IHTMLElementCollection   *pAllElement=NULL;
    CString strJsFuc = "checkForm()";
	
	if (!(pHTMLDocument = (IHTMLDocument2*)m_ctrlWeb.GetDocument()))   //获取 IHTMLDocument2 的接口指针
		return -1;
    
	pHTMLDocument->get_all(&pAllElement); 
	CComPtr<IDispatch> pDispCommon;
	CComQIPtr<IHTMLInputTextElement,&IID_IHTMLInputTextElement>pElement;
	HRESULT hr=pAllElement->item(COleVariant("username"),COleVariant((long)0),&pDispCommon);
	
	if ((hr == S_OK) && (pDispCommon != NULL))
	{
        pElement=pDispCommon;
        CString str1=Number;
        pElement->put_value(str1.AllocSysString());
        pDispCommon.Release();
		
		hr=pAllElement->item(COleVariant("passwd"),COleVariant((long)0),&pDispCommon);
		
        pElement=pDispCommon;
        CString str2="123456";
        pElement->put_value(str2.AllocSysString());
        pDispCommon.Release();
		
		
		IHTMLWindow2*  pWindow;  
		pHTMLDocument->get_parentWindow(&pWindow);
		
		VARIANT  ret;  
		
		ret.vt    =  VT_EMPTY; 
		
		BSTR bstrCode = strJsFuc.AllocSysString();
		BSTR bstrLanguage = SysAllocString(L"javascript");
		pWindow->execScript(bstrCode, bstrLanguage, &ret);
       
		return 1;
	}
	else
	{
		//ShellExecute(NULL, "open", "http://user.qzone.qq.com/316397938/main",NULL, NULL, SW_SHOWNORMAL); 
		return 2;
	}
	

}

BEGIN_EVENTSINK_MAP(CTestDlg, CDialog)
    //{{AFX_EVENTSINK_MAP(CTestDlg)
	ON_EVENT(CTestDlg, IDC_EXPLORER1, 252 /* NavigateComplete2 */, OnNavigateComplete2Explorer1, VTS_DISPATCH VTS_PVARIANT)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()

void CTestDlg::OnNavigateComplete2Explorer1(LPDISPATCH pDisp, VARIANT FAR* URL) 
{
     IHTMLDocument2   *objDocument=NULL;
        IHTMLWindow2* pIHTMLWindow = NULL;
        objDocument=(IHTMLDocument2 *)m_ctrlWeb.GetDocument();
        if(objDocument)
        {
                objDocument->get_parentWindow(&pIHTMLWindow);
                if(pIHTMLWindow)
                {
                        CString        js_str="window.alert=null;window.confirm=null;window.open = null;window.showModalDialog = null;window.onerror=function(){return true}";//这段js代码是禁止弹出一些对话框以及容错的
                        VARIANT        pvarRet;
                        pIHTMLWindow->execScript(CComBSTR(js_str), CComBSTR("JavaScript"), &pvarRet);
                        pIHTMLWindow->Release();
                }
                objDocument->Release();
        }        
	
}

void CTestDlg::OnTimer(UINT nIDEvent) 
{
	CString lastSeven_Number;
	int lastseven_number;
	int flag=0;
	// TODO: Add your control notification handler code here

     flag = FillBlank(m_CurrentNumber);
	
    if(flag==1){
	lastSeven_Number=m_CurrentNumber.Right(6);
	lastseven_number=atoi(lastSeven_Number);
	lastseven_number++; 
	lastSeven_Number.Format("%d",lastseven_number);
	m_CurrentNumber=m_CurrentNumber.Left(5)+lastSeven_Number;   
	UpdateData(FALSE);	
	}else if(flag==2)
	{
		KillTimer(1);
	}

	
	CDialog::OnTimer(nIDEvent);
}
