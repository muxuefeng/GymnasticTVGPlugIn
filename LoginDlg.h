#pragma once


class CLoginDlg : public CDialog
{
	DECLARE_DYNAMIC(CLoginDlg)

public:
	CLoginDlg(CString strServer, CWnd* pParent = NULL);  
	virtual ~CLoginDlg();

	enum { IDD = IDD_LOGIN };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    

	DECLARE_MESSAGE_MAP()
public:
	CString m_strServer;
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
};
