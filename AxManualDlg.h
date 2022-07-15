// 功能描述：体操插件
// 作者：穆学峰
// MU XUEFENG
// BEIJING CDV SPORTS

#pragma once

#include "AsIDataManager2.0.h"
#include "AsICmdTrigger2.0.h"
#include "resource.h"
#include <map>
//#include "CListCtrlEx.h"
#include "colorlistctrl.h"

typedef struct 
{
	CString bib;
	BOOL hasResult1;
	BOOL hasResult2;
	int actCount;
} player_status;

class CPlugin;
class CAxManualDlg : public CMxDialog
{
	DECLARE_DYNAMIC(CAxManualDlg)
	Mx_Declare_Dlg_ResHandle(AsDcPluginGetThisModule())

public:
	CAxManualDlg(CWnd* pParent = NULL);
	virtual ~CAxManualDlg();
	enum { IDD = IDD_Gym };

	virtual BOOL OnInitDialog();
	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BOOL Create(CWnd* pParent, CPlugin* p);
	IAsDataManager* DataM();
	IAsCmdTrigger* CmdTrigger();

public:
	void SetButtonText();
	void AdjustColumnWidth(CColorListCtrl* pList);

	void GetPlayerList(int iMatchIndex);
	void GetJudgeList();
	void GetCurPlayer();
	player_status GetPlayerScore(int matchid, CString strBib, CString lanCode);
	void AutoTriggerCmd();
	int GetPlayerAllRoundScore();
	int GetTeamScore();
	void GetEventMedal();
	void resetMatchList();
	void setLogo(int iRow, CString strAppaCode);
	void clearLogo();
	CString GetTriggerCmd(CString strRowName);
	void triggerStartList(int iCurMatch, CString cmdName);

	// datasource is Access file, then use the following function to deal
	// it is very easy!
	// db common: connect, execute sql, get field value
	BOOL ConnectDB(CString strConnection, _ConnectionPtr& pConnection);
	BOOL ExecSQL(CString strSQLString, _RecordsetPtr& pRecordSet);
	CString GetFieldValue(_RecordsetPtr pRecordSet, int iFieldIndex);
	CString GetFieldValue(_RecordsetPtr pRecordSet, CString strFieldName);
	CString getDBConnection();
	BOOL NetIsBroken();
	BOOL IsConnected();
	CString Int2CString(int ii);
	void Param2DataCenter();
	CString to_Upper(CString str);
	void Split(CString source, CStringArray& dest, CString division);
	CString m_dbString;

protected:
	virtual void DoDataExchange(CDataExchange* pDX);
	virtual void OnOK();
	virtual void OnCancel();
	afx_msg void OnComboRoundSelectChange();
	afx_msg void OnComboFieldSelectChange();
	afx_msg void OnComboStageSelectChange();
	afx_msg void OnClickListPlayer(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnNMCustomDrawList(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnClickAllRoundEvent(NMHDR* pNMHDR, LRESULT* pResult);
	DECLARE_MESSAGE_MAP()

public:

	CComboBox* m_pComboDate;
	CComboBox* m_pComboSubdivition;
	CComboBox* m_pComboRotation;

	map<int, int> m_dateKey;
	map<int, int> m_subdivisionKey;
	map<int, int> m_RotationKey;
	map<int, int> m_MatchKey;
	map<int, CString> m_groupIDKey;
	map<int, CString> m_PlayerKey;

	map<int, CString> m_stageIDKey;

	int m_dateID;
	int m_stage;
	int m_subdivisionID;
	int m_genderID;
	int m_rotationID;
	int m_MatchID;
	int m_curMatchIndex;
	int m_curPlayerIndex;
	CString m_strAppaCode;
	CString m_strAppaName;
	CString m_strTeamCode;
	CString m_strTeamName;
	CString m_PlayerBib;

	CString m_strRotationNo;
	CString m_strSubdivisionName;
	CString m_strSubdivisionNo;
	CString m_strGenderName;
	CString m_strTeamID;
	CString m_strCategoryName;

	CString m_strSubdivisionNameChinese;
	CString m_strSubdivisionNoChinese;

	map<int, CString> m_AppaNameKey;
	map<int, CString> m_AppaCodeKey;
	map<int, CString> m_RotationNoKey;
	map<int, CString> m_TeamNameKey;
	map<int, CString> m_TeamNameBibKey;
	map<int, CString> m_teamIDKey;
	map<int, CString> m_teamIDBibKey;

	map<int, CString> m_TeamCodeKey;
	map<int, CString> m_SubdivitionNameKey;
	map<int, CString> m_SubdivitionNoKey;
	map<int, CString> m_GenderNameKey;
	map<int, CString> m_GenderIDKey;
	map<int, CString> m_categoryNameKey;

	CString m_strLanguageCode;
	
	CString m_sql_get_date;
	CString m_sql_getsubdivision;
	CString m_sql_get_subdivision_bydate;
	CString m_sql_get_rotations_bysubdivision;
	CString m_sql_get_matchlist;
	CString m_sql_get_startlist;
	CString m_sql_get_playerscore;
	CString m_sql_get_player_allround_score;
	CString m_sql_get_team_score;
	CString m_sql_get_event_medal;

	CString m_strCurPlayerBib;
	CColorListCtrl m_ctrlPlayer[6];
	CColorListCtrl m_ctrlEventMedal;
//	CImageList imageList;
	CFont m_UserFont;
	CFont m_PlayerFont;
	BOOL m_bCurField;
	CMxCtrlSkin m_skin;
	CPlugin* m_pPluginImp;
	CString m_strField;
	CString m_strRound;
	int m_iEvent;// 1,4,2,3 资格赛，团体，全能，单项
	CString m_strEvent;
	int m_iSex;
	CString m_strSexDes;
	BOOL m_bChinese;

	afx_msg void OnLvnItemchangedListPlayer(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnItemchangedListEvent(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedButtonRefresh();
	afx_msg void OnBnClickedButtonGetPlayerScore();
	afx_msg void OnBnClickedCheckAd();
	afx_msg void OnBnClickedButtonTriggerTeamPreview11();
	afx_msg void OnBnClickedButtonTriggerTeamPreview12();
	afx_msg void OnBnClickedButtonTriggerTeamPreview21();
	afx_msg void OnBnClickedButtonTriggerTeamPreview22();
	afx_msg void OnBnClickedButtonTriggerTeamPreview31();
	afx_msg void OnBnClickedButtonTriggerTeamPreview32();
	afx_msg void OnBnClickedButtonTriggerTeamPreview41();
	afx_msg void OnBnClickedButtonTriggerTeamPreview42();
	afx_msg void OnBnClickedButtonTriggerTeamPreview51();
	afx_msg void OnBnClickedButtonTriggerTeamPreview52();
	afx_msg void OnBnClickedButtonTriggerTeamPreview61();
	afx_msg void OnBnClickedButtonTriggerTeamPreview62();
	afx_msg void OnBnClickedButtonGetTeamBuild();
	afx_msg void OnBnClickedButtonGetAllroundBuild();
};
