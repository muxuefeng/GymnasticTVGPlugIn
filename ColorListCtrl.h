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
	void SetItemColor(DWORD iItem, COLORREF TextColor, COLORREF TextBkColor);   //设置某一行的前景色和背景色
	void SetAllItemColor(DWORD iItem, COLORREF TextColor, COLORREF TextBkColor);//设置全部行的前景色和背景色  
	void ClearColor();                                                          //清除颜色映射表 
	CMap<DWORD, DWORD&, TEXT_BK, TEXT_BK&> MapItemColor;  
protected:
	//{{AFX_MSG(CColorListCtrl)
	//}}AFX_MSG  
	void CColorListCtrl::OnNMCustomdraw(NMHDR *pNMHDR, LRESULT *pResult); 
	DECLARE_MESSAGE_MAP() 
};
