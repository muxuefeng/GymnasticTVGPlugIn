#pragma once
#include "afxcmn.h"
typedef struct    
{  
	COLORREF colText;  
	COLORREF colTextBk;  
}TEXT_BK; 

class CColorListCtrl :public CListCtrl
{
public:
	CColorListCtrl();
	~CColorListCtrl();
public:
	void SetItemColor(DWORD iItem, COLORREF TextColor, COLORREF TextBkColor);   //����ĳһ�е�ǰ��ɫ�ͱ���ɫ
	void SetAllItemColor(DWORD iItem, COLORREF TextColor, COLORREF TextBkColor);//����ȫ���е�ǰ��ɫ�ͱ���ɫ  
	void ClearColor();                                                          //�����ɫӳ��� 
	CMap<DWORD, DWORD&, TEXT_BK, TEXT_BK&> MapItemColor;  
protected:
	//{{AFX_MSG(CColorListCtrl)
	//}}AFX_MSG  
	void CColorListCtrl::OnNMCustomdraw(NMHDR *pNMHDR, LRESULT *pResult); 
	DECLARE_MESSAGE_MAP() 
};
