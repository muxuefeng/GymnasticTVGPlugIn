#include "StdAfx.h"
#include "AsPlugin.h"

CPlugin::CPlugin(void)
{
	m_pDataManager = NULL;
	m_pCmdTrigger = NULL;
}

CPlugin::~CPlugin(void)
{
}

BOOL CPlugin::Initialize(void* pDataManager, void* pCmdTrigger, const CString& strCFGFilePath)
{
	m_pDataManager = (IAsDataManager*)pDataManager;
	m_pCmdTrigger = (IAsCmdTrigger*)pCmdTrigger;

	return TRUE;
}

BOOL CPlugin::UnInitialize()
{
	return TRUE;
}


BOOL CPlugin::Setting(CWnd* pParentWnd)
{
	if(m_dlgManualUI.GetSafeHwnd() == NULL || m_pDataManager==NULL)
		return FALSE;

	CString strServer;

	CString strTable = _T("GA_Config");
	if(m_pDataManager->TableIsExist(strTable))
	{
		strServer = m_pDataManager->GetValue(strTable, _T("P_ConnectString"), 0);
		CString strLanguageCode = m_pDataManager->GetValue(strTable, _T("P_Lang"), 0);
		m_dlgManualUI.m_strLanguageCode = strLanguageCode;

	}

	//CLoginDlg dlg(strServer, pParentWnd);
	//if(dlg.DoModal() == IDOK)
	//{
	//	strServer = dlg.m_strServer;

	//	// 2019. MU XUEFENG
	//	// 放置连接字符串到表里（数据中心）

	//	_ConnectionPtr pConnection;
	//	if(m_dlgManualUI.ConnectDB(strServer, pConnection))
	//	{
	//		m_pDataManager->SetValue(strTable, _T("P_ConnectString"), 0, strServer);

	//		AfxMessageBox(_T("数据库连接成功"));
	//	}
	//	else
	//	{
	//		AfxMessageBox(_T("数据库连接失败"));
	//	}
	//}

	return TRUE;
}

BOOL CPlugin::ManualControlUI(CWnd* pParentWnd)
{
	if ( NULL == m_dlgManualUI.GetSafeHwnd() )
	{
		m_dlgManualUI.Create(pParentWnd, this);
		//m_dlgManualUI.m_pPluginImp = this;
	}

	m_dlgManualUI.ShowWindow(SW_SHOW);	

	return TRUE;
}

BOOL CPlugin::Update()
{
	return TRUE;
}

