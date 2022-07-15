#pragma once
#include "AsIPluginDataConvertor2.0.h"
#include "AsIDataManager2.0.h"
#include "AsICmdTrigger2.0.h"
#include "AxManualDlg.h"

extern "C" __declspec(dllexport) BOOL AsQuerySpecialConvertor(void **ppConvertor);

class CPlugin : public IAsDataConvertor
{
public:
	CPlugin();
	~CPlugin();

	virtual BOOL Initialize(void* pDataManager, void* pCmdTrigger, const CString& strCFGFilePath);
	virtual BOOL UnInitialize();
	virtual BOOL Setting(CWnd* pParentWnd);
	virtual BOOL ManualControlUI(CWnd* pParentWnd);
	virtual BOOL Update();

public:
	IAsDataManager* m_pDataManager;
	IAsCmdTrigger* m_pCmdTrigger;

	CAxManualDlg m_dlgManualUI;
};
