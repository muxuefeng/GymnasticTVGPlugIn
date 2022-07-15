
#include "stdafx.h"
#include "AsPlugin.h"
#include ".\axmanualdlg.h"
#include <ctype.h>
#include <MSTCPiP.h>

// MU XUEFENG
// BEIJING CDV SPORTS

// 在单项决赛中，点左边的列表选择男女

IMPLEMENT_DYNAMIC(CAxManualDlg, CMxDialog)
CAxManualDlg::CAxManualDlg(CWnd* pParent /*=NULL*/)
	: CMxDialog(CAxManualDlg::IDD, pParent)
{
	m_pPluginImp = NULL;
	m_strField = _T("1");
	m_strRound = _T("0");
	m_iEvent = 0;
	m_bChinese = TRUE;
	m_bCurField = FALSE;


	m_pComboDate = NULL;
	m_pComboSubdivition = NULL;
	m_pComboRotation = NULL;


	m_subdivisionID = -1;
	m_rotationID = -1;
	m_MatchID = -1;
	m_PlayerBib = _T("");

	m_sql_get_date = _T("Proc_GA_TVG_PlugIn_GetDisciplineDates");
	m_sql_getsubdivision = _T("Proc_GA_TVG_PlugIn_GetSubDivision");
	m_sql_get_subdivision_bydate = _T("Proc_GA_TVG_PlugIn_GetSubDivisions_ByDate");
	m_sql_get_rotations_bysubdivision = _T("Proc_GA_TVG_PlugIn_GetRotations_BySubDivision");
	m_sql_get_matchlist = _T("Proc_GA_TVG_PlugIn_GetMatchList");
	m_sql_get_startlist = _T("Proc_GA_TVG_PlugIn_GetMatchStartList_New");
	m_sql_get_playerscore = _T("Proc_GA_TVG_PlugIn_GetMatchStartList_New");
	m_sql_get_player_allround_score = _T("Proc_GA_TVG_GetOneResultDetail");
	m_sql_get_team_score = _T("Proc_GA_TVG_GetOneTeamResults");
	m_sql_get_event_medal = _T("Proc_GA_TVG_GetEventList");

}

CAxManualDlg::~CAxManualDlg()
{
	m_UserFont.DeleteObject();
	m_PlayerFont.DeleteObject();
}

void CAxManualDlg::DoDataExchange(CDataExchange* pDX)
{
	CMxDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_PLAYER1, m_ctrlPlayer[0]);
	DDX_Control(pDX, IDC_LIST_PLAYER2, m_ctrlPlayer[1]);
	DDX_Control(pDX, IDC_LIST_PLAYER3, m_ctrlPlayer[2]);
	DDX_Control(pDX, IDC_LIST_PLAYER4, m_ctrlPlayer[3]);
	DDX_Control(pDX, IDC_LIST_PLAYER5, m_ctrlPlayer[4]);
	DDX_Control(pDX, IDC_LIST_PLAYER6, m_ctrlPlayer[5]);
	DDX_Control(pDX, IDC_LIST_ALLROUD, m_ctrlEventMedal);
}


BEGIN_MESSAGE_MAP(CAxManualDlg, CMxDialog)
	ON_CBN_SELCHANGE(IDC_COMBO_STAGE, OnComboStageSelectChange)
	ON_CBN_SELCHANGE(IDC_COMBO_FIELD, OnComboFieldSelectChange)
	ON_CBN_SELCHANGE(IDC_COMBO_ROUND, OnComboRoundSelectChange)
	ON_NOTIFY(NM_CLICK, IDC_LIST_PLAYER1, OnClickListPlayer)
	ON_NOTIFY(NM_CLICK, IDC_LIST_PLAYER2, OnClickListPlayer)
	ON_NOTIFY(NM_CLICK, IDC_LIST_PLAYER3, OnClickListPlayer)
	ON_NOTIFY(NM_CLICK, IDC_LIST_PLAYER4, OnClickListPlayer)
	ON_NOTIFY(NM_CLICK, IDC_LIST_PLAYER5, OnClickListPlayer)
	ON_NOTIFY(NM_CLICK, IDC_LIST_PLAYER6, OnClickListPlayer)
	ON_NOTIFY(NM_CLICK, IDC_LIST_ALLROUD, OnClickAllRoundEvent)
	//ON_NOTIFY(NM_CUSTOMDRAW, IDC_LIST_PLAYER1, OnNMCustomDrawList)
	//ON_NOTIFY(NM_CUSTOMDRAW, IDC_LIST_PLAYER2, OnNMCustomDrawList)
	//ON_NOTIFY(NM_CUSTOMDRAW, IDC_LIST_PLAYER3, OnNMCustomDrawList)
	//ON_NOTIFY(NM_CUSTOMDRAW, IDC_LIST_PLAYER4, OnNMCustomDrawList)
	//ON_NOTIFY(NM_CUSTOMDRAW, IDC_LIST_PLAYER5, OnNMCustomDrawList)
	//ON_NOTIFY(NM_CUSTOMDRAW, IDC_LIST_PLAYER6, OnNMCustomDrawList)
	ON_BN_CLICKED(IDC_BUTTON_REFRESH, OnBnClickedButtonRefresh)
	ON_BN_CLICKED(IDC_CHECK_AD, OnBnClickedCheckAd)
	ON_BN_CLICKED(IDC_BUTTON_GET_PLAYER_SCORE, OnBnClickedButtonGetPlayerScore)
	ON_BN_CLICKED(IDC_BUTTON_TRIGGER_STARTLIST11, OnBnClickedButtonTriggerTeamPreview11)
	ON_BN_CLICKED(IDC_BUTTON_TRIGGER_STARTLIST12, OnBnClickedButtonTriggerTeamPreview12)
	ON_BN_CLICKED(IDC_BUTTON_TRIGGER_STARTLIST21, OnBnClickedButtonTriggerTeamPreview21)
	ON_BN_CLICKED(IDC_BUTTON_TRIGGER_STARTLIST22, OnBnClickedButtonTriggerTeamPreview22)
	ON_BN_CLICKED(IDC_BUTTON_TRIGGER_STARTLIST31, OnBnClickedButtonTriggerTeamPreview31)
	ON_BN_CLICKED(IDC_BUTTON_TRIGGER_STARTLIST32, OnBnClickedButtonTriggerTeamPreview32)
	ON_BN_CLICKED(IDC_BUTTON_TRIGGER_STARTLIST41, OnBnClickedButtonTriggerTeamPreview41)
	ON_BN_CLICKED(IDC_BUTTON_TRIGGER_STARTLIST42, OnBnClickedButtonTriggerTeamPreview42)
	ON_BN_CLICKED(IDC_BUTTON_TRIGGER_STARTLIST51, OnBnClickedButtonTriggerTeamPreview51)
	ON_BN_CLICKED(IDC_BUTTON_TRIGGER_STARTLIST52, OnBnClickedButtonTriggerTeamPreview52)
	ON_BN_CLICKED(IDC_BUTTON_TRIGGER_STARTLIST61, OnBnClickedButtonTriggerTeamPreview61)
	ON_BN_CLICKED(IDC_BUTTON_TRIGGER_STARTLIST62, OnBnClickedButtonTriggerTeamPreview62)
	ON_BN_CLICKED(IDC_BUTTON_GET_TEAM_BUILD, OnBnClickedButtonGetTeamBuild)
	ON_BN_CLICKED(IDC_BUTTON_GET_ALLROUND_BUILD, OnBnClickedButtonGetAllroundBuild)
END_MESSAGE_MAP()

IAsDataManager* CAxManualDlg::DataM()
{
	return m_pPluginImp->m_pDataManager;
}

IAsCmdTrigger* CAxManualDlg::CmdTrigger()
{
	return m_pPluginImp->m_pCmdTrigger;
}

BOOL CAxManualDlg::Create(CWnd* pParent, CPlugin* p)
{
	m_pPluginImp = p;
	return CMxDialog::Create(IDD, pParent);
}

void CAxManualDlg::OnOK()
{
}

void CAxManualDlg::OnCancel()
{
}


BOOL CAxManualDlg::OnInitDialog()
{
	CMxDialog::OnInitDialog();

	m_skin.SubClassWindowChildItem(m_hWnd);

	if ( !m_pPluginImp )
		return FALSE;

	LOGFONT lf;
	memset(&lf, 0, sizeof(LOGFONT));
	lf.lfHeight = 20;
	wcscpy(lf.lfFaceName,_T("雅黑"));
	m_UserFont.CreateFontIndirect(&lf);

	LOGFONT lf2;
	memset(&lf2, 0, sizeof(LOGFONT));
	lf2.lfHeight = 30;
	lf2.lfWeight = 700;
	wcscpy(lf2.lfFaceName,_T("雅黑"));
	m_PlayerFont.CreateFontIndirect(&lf2);


	m_pComboDate = (CComboBox*)GetDlgItem(IDC_COMBO_STAGE);
	m_pComboSubdivition = (CComboBox*)GetDlgItem(IDC_COMBO_FIELD);
	m_pComboRotation = (CComboBox*)GetDlgItem(IDC_COMBO_ROUND);

	for(int i=0;i<6;i++)
	{
		m_ctrlPlayer[i].ShowWindow(FALSE);
	}


	 return TRUE;
}

BOOL CAxManualDlg::PreTranslateMessage(MSG* pMsg)
{
	if (pMsg->message == WM_KEYDOWN)
	{
		if ( m_pPluginImp )
		{
			switch (pMsg->wParam)
			{
			case 'q':
			case 'Q':
				if ( !m_pPluginImp )
				{
					break;
				}
				m_pPluginImp->QKeyDown();
				return TRUE;
			case VK_SPACE:
				if ( !m_pPluginImp )
				{
					break;
				}
				m_pPluginImp->SpaceKeyDown();

				return TRUE;
			case VK_RETURN:
			case VK_ESCAPE:
				return TRUE;
			default:
				break;

			}
		}
	}

	return CMxDialog::PreTranslateMessage(pMsg);
}

// on change Date
void CAxManualDlg::OnComboStageSelectChange()
{
	// get subdivision

	CString strDate;


	int icurSel = m_pComboDate->GetCurSel();
	if(icurSel<0)
		return;

	m_pComboDate->GetLBText(icurSel, strDate);

	if(m_dateKey.count(icurSel) > 0)
	{
		m_dateID = m_dateKey[icurSel];
	}


	CString strSql = _T("");
	strSql.Format(_T("%s '%s', '%s'"), m_sql_get_subdivision_bydate, strDate, m_strLanguageCode);
	_RecordsetPtr pRecordSet;
	if(!ExecSQL(strSql, pRecordSet))
		return;



	// Add
	m_pComboSubdivition->ResetContent();

	m_subdivisionKey.clear();
	m_stageIDKey.clear();
	m_SubdivitionNoKey.clear();
	m_SubdivitionNameKey.clear();
	m_categoryNameKey.clear();


	int iIndex = 0;
	while(!pRecordSet->adoEOF)
	{
		CString subID = GetFieldValue(pRecordSet, _T("SubdivisionID"));
		CString strSubValue = GetFieldValue(pRecordSet, _T("SubdivisionDes"));
		CString strSubNo = GetFieldValue(pRecordSet, _T("Subdivision"));

		CString strStageID = GetFieldValue(pRecordSet, _T("StageID"));
		CString strStageDes = GetFieldValue(pRecordSet, _T("StageDes"));

		CString strCategoryName = GetFieldValue(pRecordSet, _T("CategoryShortName"));

		m_subdivisionKey.insert(make_pair(iIndex, _ttoi(subID)));
		m_stageIDKey.insert(make_pair(iIndex, strStageID) );

		m_SubdivitionNoKey.insert(make_pair(iIndex, strSubNo) );
		m_SubdivitionNameKey.insert(make_pair(iIndex, strSubValue) );

		m_categoryNameKey.insert(make_pair(iIndex, strCategoryName));

		m_stageIDKey.insert(make_pair(iIndex, strStageID) );

		m_pComboSubdivition->AddString(strSubValue);

		pRecordSet->MoveNext();
		iIndex++;
	}
	
}

void CAxManualDlg::OnComboFieldSelectChange()
{

	int icurSel = m_pComboSubdivition->GetCurSel();
	if(icurSel<0)
		return;

	m_pComboRotation->ResetContent();

	m_RotationKey.clear();
	m_RotationNoKey.clear();

	CString strSubValue;
	m_pComboSubdivition->GetLBText(icurSel, strSubValue);

	map<int, int>::iterator it;
	it = m_subdivisionKey.find(icurSel);
	if(it != m_subdivisionKey.end())
	{
		m_subdivisionID = it->second;
	}
	else
	{
		return;
	}


	if(m_subdivisionKey.count(icurSel) > 0)
	{
		m_strSubdivisionNo = m_SubdivitionNoKey[icurSel];
		m_strSubdivisionName = m_SubdivitionNameKey[icurSel];
		m_stage = _ttoi(m_stageIDKey[icurSel]);

		m_strCategoryName = m_categoryNameKey[icurSel];
	}

	CString strSQLString;
	strSQLString.Format(_T("exec %s '%d'"), m_sql_get_rotations_bysubdivision, m_subdivisionID);

	if(strSQLString == _T("")) return;

	_RecordsetPtr pRecordSet;
	if(!ExecSQL(strSQLString, pRecordSet))
		return;

	int iIndex = 0;
	while(!pRecordSet->adoEOF)
	{
		CString strRotation = GetFieldValue(pRecordSet, _T("Rotation") );
		CString strRotatonID = GetFieldValue(pRecordSet, _T("F_RotationID"));

		m_RotationKey.insert(make_pair(iIndex, _ttoi(strRotatonID)));
		m_RotationNoKey.insert(make_pair(iIndex, strRotation));

		m_pComboRotation->AddString(strRotation);


		pRecordSet->MoveNext();
		iIndex++;
	}
	

	Param2DataCenter();
}

void CAxManualDlg::OnComboRoundSelectChange()
{

	GetDlgItem(IDC_BUTTON_GET_PLAYER_SCORE)->ShowWindow(TRUE);

	if(m_stage == 1)
	{
		GetDlgItem(IDC_BUTTON_GET_TEAM_BUILD)->ShowWindow(TRUE);
		GetDlgItem(IDC_BUTTON_GET_ALLROUND_BUILD)->ShowWindow(TRUE);
	}
	if(m_stage == 4)
	{
		GetDlgItem(IDC_BUTTON_GET_TEAM_BUILD)->ShowWindow(TRUE);
		GetDlgItem(IDC_BUTTON_GET_ALLROUND_BUILD)->ShowWindow(FALSE);
	}
	if(m_stage == 3)
	{
		GetDlgItem(IDC_BUTTON_GET_TEAM_BUILD)->ShowWindow(FALSE);
		GetDlgItem(IDC_BUTTON_GET_ALLROUND_BUILD)->ShowWindow(FALSE);
	}
	if(m_stage == 2)
	{
		GetDlgItem(IDC_BUTTON_GET_TEAM_BUILD)->ShowWindow(FALSE);
		GetDlgItem(IDC_BUTTON_GET_ALLROUND_BUILD)->ShowWindow(TRUE);
	}
	int icurSelIndex = m_pComboRotation->GetCurSel();
	if(icurSelIndex < 0)
		return;


	CString strMyRound;
	m_pComboRotation->GetLBText(icurSelIndex, strMyRound);


	map<int, int>::iterator it;
	it = m_RotationKey.find(icurSelIndex);
	if(it != m_RotationKey.end())
	{
		int irotiD2 = it->second;
		m_rotationID = irotiD2;
	}


	if(m_RotationKey.count(icurSelIndex) > 0)
	{
		m_strRotationNo = m_RotationNoKey[icurSelIndex];
		m_rotationID = m_RotationKey[icurSelIndex];

	}


	CString strSQLString;
	strSQLString.Format(_T("%s %d, %d, '%s'"), m_sql_get_matchlist, m_subdivisionID, m_rotationID, m_strLanguageCode);

	if(strSQLString == _T("")) return;


	_RecordsetPtr pRecordSet;
	if(!ExecSQL(strSQLString, pRecordSet))
		return;


	m_MatchKey.clear();
	m_groupIDKey.clear();
	m_AppaNameKey.clear();
	m_AppaCodeKey.clear();
	m_GenderIDKey.clear();
	m_GenderNameKey.clear();


	clearLogo();

	for(int i=0;i<6;i++)
	{
		GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST11)->ShowWindow(FALSE);
		GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST12)->ShowWindow(FALSE);
		GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST21)->ShowWindow(FALSE);
		GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST22)->ShowWindow(FALSE);
		GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST31)->ShowWindow(FALSE);
		GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST32)->ShowWindow(FALSE);
		GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST41)->ShowWindow(FALSE);
		GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST42)->ShowWindow(FALSE);
		GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST51)->ShowWindow(FALSE);
		GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST52)->ShowWindow(FALSE);
		GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST61)->ShowWindow(FALSE);
		GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST62)->ShowWindow(FALSE);
	}

	int iRow = 0;
	while(!pRecordSet->adoEOF)
	{

		CString strAppaName = GetFieldValue(pRecordSet, _T("ApparatusName"));
		CString strAppaCode = GetFieldValue(pRecordSet, _T("ApparatusCode"));
		CString strMatchID = GetFieldValue(pRecordSet, _T("F_MatchID"));
		CString strGroupName = GetFieldValue(pRecordSet, _T("GroupName"));
		CString strSexCode = GetFieldValue(pRecordSet, _T("F_SexCode"));
		CString strSexName = GetFieldValue(pRecordSet, _T("F_SexLongName"));

		setLogo(iRow, strAppaCode);

		if(iRow == 0)
		{
			m_MatchID = _ttoi(strMatchID);
		}

		if(iRow == 0)
		{
			GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST11)->ShowWindow(TRUE);
			GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST12)->ShowWindow(TRUE);
		}
		if(iRow == 1)
		{
			GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST21)->ShowWindow(TRUE);
			GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST22)->ShowWindow(TRUE);
		}
		if(iRow == 2)
		{
			GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST31)->ShowWindow(TRUE);
			GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST32)->ShowWindow(TRUE);
		}
		if(iRow == 3)
		{
			GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST41)->ShowWindow(TRUE);
			GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST42)->ShowWindow(TRUE);
		}
		if(iRow == 4)
		{
			GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST51)->ShowWindow(TRUE);
			GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST52)->ShowWindow(TRUE);
		}
		if(iRow == 5)
		{
			GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST61)->ShowWindow(TRUE);
			GetDlgItem(IDC_BUTTON_TRIGGER_STARTLIST62)->ShowWindow(TRUE);
		}

		m_MatchKey.insert(make_pair(iRow, _ttoi(strMatchID)));
		m_AppaNameKey.insert(make_pair(iRow, strAppaName));
		m_AppaCodeKey.insert(make_pair(iRow, strAppaCode));
		m_groupIDKey.insert(make_pair(iRow, strGroupName));
		m_GenderIDKey.insert(make_pair(iRow, strSexCode));
		m_GenderNameKey.insert(make_pair(iRow, strSexName));
		

		pRecordSet->MoveNext();
		iRow++;
	}
	m_skin.SubClassWindowChildItem(m_hWnd);

	resetMatchList();
	Param2DataCenter();
}

void CAxManualDlg::clearLogo()
{
	CStatic *pStatic1=(CStatic*)GetDlgItem(IDC_STATIC_LOGO1);
	pStatic1-> ModifyStyle(0x0,SS_ICON|SS_CENTERIMAGE); 
	pStatic1-> SetIcon(NULL); 

	pStatic1=(CStatic*)GetDlgItem(IDC_STATIC_LOGO2);
	pStatic1-> ModifyStyle(0x0,SS_ICON|SS_CENTERIMAGE); 
	pStatic1-> SetIcon(NULL); 

	pStatic1=(CStatic*)GetDlgItem(IDC_STATIC_LOGO3);
	pStatic1-> ModifyStyle(0x0,SS_ICON|SS_CENTERIMAGE); 
	pStatic1-> SetIcon(NULL); 

	pStatic1=(CStatic*)GetDlgItem(IDC_STATIC_LOGO4);
	pStatic1-> ModifyStyle(0x0,SS_ICON|SS_CENTERIMAGE); 
	pStatic1-> SetIcon(NULL); 

	pStatic1=(CStatic*)GetDlgItem(IDC_STATIC_LOGO5);
	pStatic1-> ModifyStyle(0x0,SS_ICON|SS_CENTERIMAGE); 
	pStatic1-> SetIcon(NULL); 

	pStatic1=(CStatic*)GetDlgItem(IDC_STATIC_LOGO6);
	pStatic1-> ModifyStyle(0x0,SS_ICON|SS_CENTERIMAGE); 
	pStatic1-> SetIcon(NULL); 
}
void CAxManualDlg::setLogo(int iRow, CString strAppaCode)
{
	if(iRow>5)
	{
		return;
	}

	CStatic *pStatic = NULL;
	if(iRow == 0)
	{
		pStatic = (CStatic*)GetDlgItem(IDC_STATIC_LOGO1);
	}
	if(iRow == 1)
	{
		pStatic = (CStatic*)GetDlgItem(IDC_STATIC_LOGO2);
	}
	if(iRow == 2)
	{
		pStatic = (CStatic*)GetDlgItem(IDC_STATIC_LOGO3);
	}
	if(iRow == 3)
	{
		pStatic = (CStatic*)GetDlgItem(IDC_STATIC_LOGO4);
	}
	if(iRow == 4)
	{
		pStatic = (CStatic*)GetDlgItem(IDC_STATIC_LOGO5);
	}
	if(iRow == 5)
	{
		pStatic = (CStatic*)GetDlgItem(IDC_STATIC_LOGO6);
	}

	pStatic-> ModifyStyle(0x0,SS_ICON|SS_CENTERIMAGE); 

	if(strAppaCode == _T("FX"))
	{
		pStatic-> SetIcon(AfxGetApp()->LoadIcon(IDI_FX)); 
	}
	if(strAppaCode == _T("PH"))
	{
		pStatic-> SetIcon(AfxGetApp()->LoadIcon(IDI_PH)); 
	}
	if(strAppaCode == _T("SR"))
	{
		pStatic-> SetIcon(AfxGetApp()->LoadIcon(IDI_R)); 
	}
	if(strAppaCode == _T("VT"))
	{
		pStatic-> SetIcon(AfxGetApp()->LoadIcon(IDI_VT)); 
	}
	if(strAppaCode == _T("PB"))
	{
		pStatic-> SetIcon(AfxGetApp()->LoadIcon(IDI_PB)); 
	}
	if(strAppaCode == _T("HB"))
	{
		pStatic-> SetIcon(AfxGetApp()->LoadIcon(IDI_HB)); 
	}
	if(strAppaCode == _T("UB"))
	{
		pStatic-> SetIcon(AfxGetApp()->LoadIcon(IDI_UB)); 
	}
	if(strAppaCode == _T("BB"))
	{
		pStatic-> SetIcon(AfxGetApp()->LoadIcon(IDI_BB)); 
	}

}

void CAxManualDlg::GetPlayerList(int iMatchIndex)
{
	if(iMatchIndex<0)
	{
		return;
	}


	CString strSQLString;
	strSQLString.Format(_T("%s  %d, '%s'"), m_sql_get_startlist, m_MatchKey[iMatchIndex], m_strLanguageCode);

	if(strSQLString == _T("")) return;

	_RecordsetPtr pRecordSet;
	if(!ExecSQL(strSQLString, pRecordSet))
		return;
	int iRow = 0;
	while(!pRecordSet->adoEOF)
	{
		CString strBib = GetFieldValue(pRecordSet, _T("F_Bib"));
		CString strPlayer = GetFieldValue(pRecordSet, _T("F_LongName"));
		CString strTeamName = GetFieldValue(pRecordSet, _T("F_TeamName"));
		CString strTeamCode = GetFieldValue(pRecordSet, _T("F_TeamCode"));//队伍Code
		CString strTeamID = GetFieldValue(pRecordSet, _T("F_DelegationID"));
		CString strScore = GetFieldValue(pRecordSet, _T("F_TotalPoints"));

		m_ctrlPlayer[iMatchIndex].InsertItem(iRow, _T(""));
		m_ctrlPlayer[iMatchIndex].SetItemText(iRow, 0, strBib );
		m_ctrlPlayer[iMatchIndex].SetItemText(iRow, 1, strPlayer );
		m_ctrlPlayer[iMatchIndex].SetItemText(iRow, 2, strTeamName );
		m_ctrlPlayer[iMatchIndex].SetItemText(iRow, 3, strScore );


		pRecordSet->MoveNext();
		iRow++;
	}


}

int CAxManualDlg::GetPlayerAllRoundScore()
{
	int iSum = 0;
	CString strSQLString;
	strSQLString.Format(_T("%s %d, '%s'"), m_sql_get_player_allround_score, m_MatchID, m_PlayerBib);

	if(strSQLString == _T("")) return 0;

	_RecordsetPtr pRecordSet;
	if(!ExecSQL(strSQLString, pRecordSet))
		return 0;
	int iRow = 0;
	while(!pRecordSet->adoEOF)
	{
		CString strScore1 = GetFieldValue(pRecordSet, _T("App_1_Score"));
		CString strScore2 = GetFieldValue(pRecordSet, _T("App_2_Score"));
		CString strScore3 = GetFieldValue(pRecordSet, _T("App_3_Score"));
		CString strScore4 = GetFieldValue(pRecordSet, _T("App_4_Score"));
		CString strScore5 = GetFieldValue(pRecordSet, _T("App_5_Score"));
		CString strScore6 = GetFieldValue(pRecordSet, _T("App_6_Score"));

		if(strScore1.GetLength()>0)
		{
			iSum++;
		}
		if(strScore2.GetLength()>0)
		{
			iSum++;
		}
		if(strScore3.GetLength()>0)
		{
			iSum++;
		}
		if(strScore4.GetLength()>0)
		{
			iSum++;
		}
		if(strScore5.GetLength()>0)
		{
			iSum++;
		}
		if(strScore6.GetLength()>0)
		{
			iSum++;
		}
		pRecordSet->MoveNext();
		iRow++;
	}

	return iSum;
}

int CAxManualDlg::GetTeamScore()
{
	int iSum = 0;
	CString strSQLString;
	strSQLString.Format(_T("%s %d, '%s'"), m_sql_get_team_score, m_MatchID, m_strTeamID);

	if(strSQLString == _T("")) return 0;

	_RecordsetPtr pRecordSet;
	if(!ExecSQL(strSQLString, pRecordSet))
		return 0;
	int iRow = 0;
	while(!pRecordSet->adoEOF)
	{
		CString strScore1 = GetFieldValue(pRecordSet, _T("App1_Score"));
		CString strScore2 = GetFieldValue(pRecordSet, _T("App2_Score"));
		CString strScore3 = GetFieldValue(pRecordSet, _T("App3_Score"));
		CString strScore4 = GetFieldValue(pRecordSet, _T("App4_Score"));
		CString strScore5 = GetFieldValue(pRecordSet, _T("App5_Score"));
		CString strScore6 = GetFieldValue(pRecordSet, _T("App6_Score"));

		if(strScore1.GetLength()>0)
		{
			iSum++;
		}
		if(strScore2.GetLength()>0)
		{
			iSum++;
		}
		if(strScore3.GetLength()>0)
		{
			iSum++;
		}
		if(strScore4.GetLength()>0)
		{
			iSum++;
		}
		if(strScore5.GetLength()>0)
		{
			iSum++;
		}
		if(strScore6.GetLength()>0)
		{
			iSum++;
		}
		pRecordSet->MoveNext();
		iRow++;
	}

	return iSum;
}

CString CAxManualDlg::GetTriggerCmd(CString strRowName)
{
	CString cmdName = _T("");

	CString strSQLString;
	strSQLString.Format(_T("SELECT cmdName from GA_Trigger WHERE fnName = '%s' "), strRowName);

	STableRecordSet pRecordSet;

	if(DataM()!=NULL)
	{
		if(DataM()->ExecuteSQL(strSQLString, &pRecordSet))
		{
			cmdName = pRecordSet.GetFieldValue(0, 0);
		}
	}


	return cmdName;
}

void CAxManualDlg::resetMatchList()
{
	for(int i=0;i<6;i++)
	{
		m_ctrlPlayer[i].ShowWindow(FALSE);
	}

	SetDlgItemText(IDC_STATIC_GROUP1, _T(""));
	SetDlgItemText(IDC_STATIC_GROUP2, _T(""));
	SetDlgItemText(IDC_STATIC_GROUP3, _T(""));
	SetDlgItemText(IDC_STATIC_GROUP4, _T(""));
	SetDlgItemText(IDC_STATIC_GROUP5, _T(""));
	SetDlgItemText(IDC_STATIC_GROUP6, _T(""));


	DWORD tdwValue=LVS_EX_GRIDLINES|LVS_EX_FULLROWSELECT ;


	for(int i=0;i<m_MatchKey.size();i++)
	{
		if(i>5)
		{
			break;
		}

		if(i==0)
		{
			SetDlgItemText(IDC_STATIC_GROUP1, m_groupIDKey[i]);
		}
		if(i==1)
		{
			SetDlgItemText(IDC_STATIC_GROUP2, m_groupIDKey[i]);
		}
		if(i==2)
		{
			SetDlgItemText(IDC_STATIC_GROUP3, m_groupIDKey[i]);
		}
		if(i==3)
		{
			SetDlgItemText(IDC_STATIC_GROUP4, m_groupIDKey[i]);
		}
		if(i==4)
		{
			SetDlgItemText(IDC_STATIC_GROUP5, m_groupIDKey[i]);
		}
		if(i==5)
		{
			SetDlgItemText(IDC_STATIC_GROUP6, m_groupIDKey[i]);
		}
		

		m_ctrlPlayer[i].ShowWindow(TRUE);
		m_ctrlPlayer[i].SetFont(&m_UserFont);
		m_ctrlPlayer[i].SetExtendedStyle (tdwValue);

		m_ctrlPlayer[i].DeleteAllItems();
		while(m_ctrlPlayer[i].DeleteColumn(0));

		int iCol = 0;
		m_ctrlPlayer[i].InsertColumn(iCol++, _T("号码"), LVCFMT_LEFT, 50);
		m_ctrlPlayer[i].InsertColumn(iCol++, _T("姓名"), LVCFMT_LEFT, 80);
		m_ctrlPlayer[i].InsertColumn(iCol++, _T("队名"), LVCFMT_LEFT, 70);
		m_ctrlPlayer[i].InsertColumn(iCol++, _T("得分"), LVCFMT_RIGHT, 120);
		GetPlayerList(i);
		//AdjustColumnWidth(&m_ctrlPlayer[i]);
	}


}

CString CAxManualDlg::to_Upper(CString str)
{
	CString strRet = str;

	for (int i = 0; i <strRet.GetLength(); i++)
	{
		int ii = strRet.GetAt(i);
		if(ii ==48)
		{
			strRet.Replace(_T("0"), _T("零"));
		}
		else if (ii == 49)
		{
			strRet.Replace(_T("1"), _T("一"));
		}
		else if (ii == 50)
		{
			strRet.Replace(_T("2"), _T("二"));
		}
		else if (ii == 51)
		{
			strRet.Replace(_T("3"), _T("三"));
		}
		else if (ii == 52)
		{
			strRet.Replace(_T("4"), _T("四"));
		}
		else if (ii == 53)
		{
			strRet.Replace(_T("5"), _T("五"));
		}
		else if (ii == 54)
		{
			strRet.Replace(_T("6"), _T("六"));
		}
		else if (ii == 55)
		{
			strRet.Replace(_T("7"), _T("七"));
		}
		else if (ii == 56)
		{
			strRet.Replace(_T("8"), _T("八"));
		}
		else if (ii == 57)
		{
			strRet.Replace(_T("9"), _T("九"));
		}
	}

	return strRet;
}

void CAxManualDlg::Split(CString source, CStringArray& dest, CString division)
{
	dest.RemoveAll();
	int pos = 0;
	int pre_pos = 0;
	while( -1 != pos )
	{
		pre_pos = pos;
		pos = source.Find(division,(pos+1));
		dest.Add(source.Mid(pre_pos,(pos-pre_pos)));
		pre_pos++;
	}
}

CString CAxManualDlg::getDBConnection()
{
	CString strServer;
	CString strDB;
	CString strUser;
	CString strPwd;

	CString strTable = _T("DSYS_DataBaseConnection");
	CString dbString = DataM()->GetValue(strTable, _T("_connectstring"), 0);

	CStringArray array;
	CAsSplitString strSplitData;
	strSplitData.SetSplitFlag(_T(";"));
	strSplitData.SetData(dbString);
	strSplitData.GetSplitStrArray(array);

	for(int i=0;i<array.GetCount();i++)
	{
		CStringArray arrayItem;
		CString item = array.GetAt(i);
		CAsSplitString strItemData;
		strItemData.SetSplitFlag(_T("="));
		strItemData.SetData(item);
		strItemData.GetSplitStrArray(arrayItem);

		if(arrayItem.GetCount()==2)
		{
			CString str1 = arrayItem.GetAt(0);
			//str1 = str1.Replace(';', '');
			CString str2 = arrayItem.GetAt(1);

			if(str1 == _T("Data Source"))
			{
				strServer = str2;
			}
			if(str1 == _T("Initial Catalog"))
			{
				strDB = str2;
			}
			if(str1 == _T("User ID"))
			{
				strUser = str2;
			}
			if(str1 == _T("Password"))
			{
				strPwd = str2;
			}
		}
	}

	CString conString;
	conString.Format(_T("Driver={SQL Server};Server=%s;	Trusted_Connection=no; Database=%s;User ID=%s; PWD=%s;Connect Timeout=5"), strServer, strDB, strUser, strPwd);

	return conString;
}

void CAxManualDlg::Param2DataCenter()
{
	CString strTable = _T("GA_Config");
	DataM()->SetValue(strTable, _T("P_DateID"), 0, Int2CString(m_dateID));
	DataM()->SetValue(strTable, _T("P_Stage"), 0, Int2CString(m_stage));
	DataM()->SetValue(strTable, _T("P_SubDivisionID"), 0, Int2CString(m_subdivisionID));
	DataM()->SetValue(strTable, _T("P_GenderID"), 0, Int2CString(m_genderID));
	DataM()->SetValue(strTable, _T("P_AppaCode"), 0, m_strAppaCode);
	DataM()->SetValue(strTable, _T("P_RotationID"), 0, Int2CString(m_rotationID));
	DataM()->SetValue(strTable, _T("P_MatchID"), 0, Int2CString(m_MatchID));
	DataM()->SetValue(strTable, _T("P_TeamCode"), 0, m_strTeamCode);
	DataM()->SetValue(strTable, _T("P_Bib"), 0, m_PlayerBib);


	DataM()->SetValue(strTable, _T("P_AppaName"), 0, m_strAppaName);
	DataM()->SetValue(strTable, _T("P_RotationNo"), 0, m_strRotationNo);
	DataM()->SetValue(strTable, _T("P_TeamName"), 0, m_strTeamName);
	DataM()->SetValue(strTable, _T("P_SubdivisionName"), 0, m_strSubdivisionName);
	DataM()->SetValue(strTable, _T("P_SubdivisionNo"), 0, m_strSubdivisionNo);
	DataM()->SetValue(strTable, _T("P_GenderName"), 0, m_strGenderName);
	DataM()->SetValue(strTable, _T("P_TeamID"), 0, m_strTeamID);
	DataM()->SetValue(strTable, _T("P_CategoryName"), 0, m_strCategoryName);

	


	m_strSubdivisionNameChinese = to_Upper(m_strSubdivisionName);
	m_strSubdivisionNoChinese = to_Upper(m_strSubdivisionNo);

	// AfxMessageBox(m_strSubdivisionNameChinese);

	DataM()->SetValue(strTable, _T("P_SubdivisionNameChinese"), 0, m_strSubdivisionNameChinese);
	DataM()->SetValue(strTable, _T("P_SubdivisionNoChinese"), 0, m_strSubdivisionNoChinese);

	

	CString strMsg;
	strMsg.Format(_T("DateID:%d, Stage:%d, SubDivisionID:%d, RotaionID:%d, Gender:%d, AppaCode:%s, MatchID:%d, Bib:%s")
		, m_dateID, m_stage, m_subdivisionID, m_rotationID, m_genderID, m_strAppaCode, m_MatchID, m_PlayerBib);
	SetDlgItemText(IDC_STATIC_MSG, strMsg);
}

CString CAxManualDlg::Int2CString(int ii)
{
	CString strRet;

	strRet.Format(_T("%d"), ii);

	return strRet;
}

void CAxManualDlg::OnClickAllRoundEvent(NMHDR* pNMHDR, LRESULT* pResult)
{
	int iSel = m_ctrlEventMedal.GetSelectionMark();
	CString strEventID = m_ctrlEventMedal.GetItemText(iSel, 0);
	CString strEventName = m_ctrlEventMedal.GetItemText(iSel, 1);
	DataM()->SetVariableValue(_T("EventID"), strEventID);
	DataM()->SetValue(_T("GA_MatchInfo"), _T("MedalEventName"), 0, strEventName);
}

void CAxManualDlg::OnClickListPlayer(NMHDR* pNMHDR, LRESULT* pResult)
{
	UINT_PTR idfrom = pNMHDR->idFrom;
	int iCurMatch = idfrom - 1002;
	if(iCurMatch<0)
	{
		return;
	}
	m_curMatchIndex = iCurMatch;

	m_strGenderName = m_GenderNameKey[m_curMatchIndex];
	m_genderID = _ttoi(m_GenderIDKey[m_curMatchIndex]);

	m_strAppaCode = m_AppaCodeKey[m_curMatchIndex];
	m_strAppaName = m_AppaNameKey[m_curMatchIndex];

	m_MatchID = m_MatchKey[iCurMatch];
	int iSel = m_ctrlPlayer[iCurMatch].GetSelectionMark();
	m_curPlayerIndex = iSel;
	m_PlayerBib = m_ctrlPlayer[iCurMatch].GetItemText(iSel, 0);

	for(int i=0;i<m_MatchKey.size();i++)
	{
		int iRows = m_ctrlPlayer[i].GetItemCount();
		for(int j=0;j<iRows;j++)
		{
			m_ctrlPlayer[i].SetItemColor(j, RGB(0,0,0),RGB(184,184,184));
		}
	}

	m_ctrlPlayer[m_curMatchIndex].SetItemState(m_curPlayerIndex, LVIS_SELECTED|LVIS_FOCUSED, LVIS_SELECTED|LVIS_FOCUSED);
	m_ctrlPlayer[m_curMatchIndex].SetSelectionMark(m_curPlayerIndex);
	m_ctrlPlayer[m_curMatchIndex].SetItemColor(m_curPlayerIndex,RGB(255,255,255),RGB(10, 36, 106));


	//m_ctrlPlayer[iCurMatch].SetItemState(iSel, LVIS_SELECTED|LVIS_ACTIVATING|LVIS_FOCUSED, LVIS_SELECTED|LVIS_FOCUSED);
	//m_ctrlPlayer[iCurMatch].SetSelectionMark(iSel);

	//DWORD tdwValue1=LVS_EX_GRIDLINES|LVS_EX_FULLROWSELECT;
	//DWORD tdwValue2=LVS_EX_GRIDLINES|LVS_EX_FULLROWSELECT|LVS_SHOWSELALWAYS ;

	//for(int i=0;i<6;i++)
	//{
	//	m_ctrlPlayer[i].SetExtendedStyle(tdwValue1);
	//	m_ctrlPlayer[i].SetSelectionMark(-1);
	//}

	//LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	//NMLVCUSTOMDRAW* pLVCD = reinterpret_cast<NMLVCUSTOMDRAW*>(pNMHDR);

	//int nItem = pNMItemActivate->iItem;    //获取选中的单元格所在行
	//int nSubItem = pNMItemActivate->iSubItem;  //获取选中的单元格所在列

	//if( m_ctrlPlayer[iCurMatch].GetItemState(nItem, CDIS_SELECTED) )
	//{
	//	COLORREF clrNewTextColor, clrNewBkColor;
	//	clrNewTextColor = RGB(255,255,255);  
	//	clrNewBkColor = RGB(49,106,197); 

	//	pLVCD->nmcd.uItemState ^= CDIS_SELECTED;
	//	pLVCD->clrText = clrNewTextColor;
	//	pLVCD->clrTextBk = clrNewBkColor;
	//}

	AutoTriggerCmd();

}

void CAxManualDlg::AutoTriggerCmd()
{
	player_status ps = GetPlayerScore(m_MatchID, m_PlayerBib, m_strLanguageCode);

	Param2DataCenter();

	if(m_stage == 2)
	{
		CString strCmd = GetTriggerCmd(_T("全能DE"));
		if(GetPlayerAllRoundScore()>1)
		{
			strCmd = GetTriggerCmd(_T("全能Build"));
		}
		else if(GetPlayerAllRoundScore()==0)
		{
			strCmd = GetTriggerCmd(_T("运动员介绍"));
		}
		CmdTrigger()->TriggerCmd(strCmd, IAsCmdTrigger::emTriggerPreview);
	}
	else
	{
		CString strCmd = GetTriggerCmd(_T("运动员介绍"));
		if(ps.actCount == 2)
		{
			if(ps.hasResult2)
			{
				strCmd = GetTriggerCmd(_T("跳马第二跳成绩"));
			}
			else
			{
				strCmd = GetTriggerCmd(_T("跳马第二跳介绍"));
			}
		}
		else
		{
			if(ps.hasResult1)
			{
				strCmd = GetTriggerCmd(_T("运动员成绩"));
			}
		}
		CmdTrigger()->TriggerCmd(strCmd, IAsCmdTrigger::emTriggerPreview);
	}
}

player_status CAxManualDlg::GetPlayerScore(int matchid, CString strBib, CString lanCode)
{
	player_status ps;
	ps.hasResult1 = FALSE;
	ps.hasResult2 = FALSE;
	ps.actCount = 0;
	ps.bib = strBib;

	CString strSQLString;
	strSQLString.Format(_T("exec %s %d, '%s', '%s' "), m_sql_get_playerscore, matchid, lanCode, strBib);

	if(strSQLString == _T("")) 
	{
		return ps;
	}

	_RecordsetPtr pRecordSet;
	if(!ExecSQL(strSQLString, pRecordSet))
	{
		return ps;
	}

	CStatic* pStaticCtrl = (CStatic*)GetDlgItem(IDC_STATIC_CURRENT_PLAYER);
	pStaticCtrl->SetFont(&m_PlayerFont);


	int iIndex = 0;
	while(!pRecordSet->adoEOF)
	{
		CString strRank = GetFieldValue(pRecordSet, _T("F_Rank") );
		CString strTotal = GetFieldValue(pRecordSet, _T("F_TotalPoints"));
		CString strTotal2 = GetFieldValue(pRecordSet, _T("F_TotalPoints_2"));
		CString strDelegationID = GetFieldValue(pRecordSet, _T("F_DelegationID"));
		CString strActCount = GetFieldValue(pRecordSet, _T("F_ActCount"));
		
		m_strTeamID = strDelegationID;

		// bib , player,  score
		pStaticCtrl->SetWindowText(_T("BIB:") + m_PlayerBib + _T(" - SCORE:") + strTotal);

		if(strTotal.GetLength()>0)
		{
			ps.hasResult1 = TRUE;
		}
		if(strTotal2.GetLength()>0)
		{
			ps.hasResult2 = TRUE;
		}

		ps.actCount = _ttoi(strActCount);

		pRecordSet->MoveNext();
		iIndex++;
	}

	return ps;
}


void CAxManualDlg::GetEventMedal()
{

	CString strSQLString;
	strSQLString.Format(_T("exec %s '%s' "), m_sql_get_event_medal, m_strLanguageCode);

	if(strSQLString == _T("")) return;

	_RecordsetPtr pRecordSet;
	if(!ExecSQL(strSQLString, pRecordSet))
		return;

	DWORD tdwValue=LVS_EX_GRIDLINES|LVS_EX_FULLROWSELECT;
	m_ctrlEventMedal.SetExtendedStyle (tdwValue);

	m_ctrlEventMedal.DeleteAllItems();
	while(m_ctrlEventMedal.DeleteColumn(0));

	int iCol = 0;
	m_ctrlEventMedal.InsertColumn(iCol++, _T("EventID"));
	m_ctrlEventMedal.InsertColumn(iCol++, _T("项目名称"));

	int iIndex = 0;
	while(!pRecordSet->adoEOF)
	{
		CString strEventID = GetFieldValue(pRecordSet, _T("F_EventID") );
		CString strEventName = GetFieldValue(pRecordSet, _T("EventName"));

		m_ctrlEventMedal.InsertItem(iIndex, _T(""));
		m_ctrlEventMedal.SetItemText(iIndex, 0, strEventID );
		m_ctrlEventMedal.SetItemText(iIndex, 1, strEventName );

		pRecordSet->MoveNext();
		iIndex++;
	}

	AdjustColumnWidth(&m_ctrlEventMedal);

}

BOOL CAxManualDlg::ConnectDB(CString strConnection, _ConnectionPtr& pConnection)
{
	HRESULT hr;
	try
	{
		hr = pConnection.CreateInstance("ADODB.Connection");

		if(SUCCEEDED(hr))
		{
			hr = pConnection->Open((_bstr_t)strConnection,"","",adModeUnknown);
			return TRUE;
		}
	}
	catch(_com_error e)
	{
		return FALSE;
	}  

	return FALSE;
}

BOOL CAxManualDlg::IsConnected()
{
	CString strTestSql = _T("select count(*) from t_team");

	_CommandPtr commandptr;
	_RecordsetPtr pRecordSet;

	BOOL bSuccess = FALSE;
	try
	{
		commandptr.CreateInstance (__uuidof(Command));
		pRecordSet.CreateInstance (__uuidof(Recordset));
//		commandptr->ActiveConnection = m_pConnection;
		bSuccess = TRUE;
	}
	catch(...)
	{
		bSuccess = FALSE;
	}

	if(!bSuccess)
		return FALSE;

	_variant_t fieldCount;
	VariantInit(&fieldCount);

	try
	{
		commandptr->CommandType =adCmdText;
		commandptr->CommandText =(_bstr_t)strTestSql;
		pRecordSet = commandptr->Execute(&fieldCount,NULL,adCmdUnknown);
		{
			if(fieldCount.vt!=VT_EMPTY)
			{
				return TRUE;
			}
		}
	}
	catch(...)
	{
		return FALSE;
	}

	return FALSE;

}

BOOL CAxManualDlg::ExecSQL(CString strSQLString, _RecordsetPtr& pRecordSet)
{
	CString strConnectSting = m_dbString;

	_ConnectionPtr m_pConnection;

	ConnectDB(strConnectSting, m_pConnection);

	_CommandPtr commandptr;

	BOOL bSuccess = FALSE;
	try
	{
		commandptr.CreateInstance (__uuidof(Command));
		pRecordSet.CreateInstance (__uuidof(Recordset));
		commandptr->ActiveConnection = m_pConnection;
		bSuccess = TRUE;
	}
	catch(...)
	{
		bSuccess = FALSE;
	}

	if(!bSuccess)
		return FALSE;

	_variant_t fieldCount;
	VariantInit(&fieldCount);

	try
	{
		commandptr->CommandType =adCmdText;
		commandptr->CommandText =(_bstr_t)strSQLString;
		pRecordSet = commandptr->Execute(&fieldCount,NULL,adCmdUnknown);
		{
			if(fieldCount.vt!=VT_EMPTY)
			{
				return TRUE;
			}
		}
	}
	catch(...)
	{
		return FALSE;
	}

	return FALSE;
}

CString CAxManualDlg::GetFieldValue(_RecordsetPtr pRecordSet, int iFieldIndex)
{
	CString strReturn = _T("");

	FieldPtr m_fieldCtl;
	int fieldCount = pRecordSet->Fields->Count;
	int fieldLength = 0;

	int nItem = 0;
	_variant_t varValue;
	_bstr_t bstrValue;

	if(iFieldIndex>fieldCount || pRecordSet==NULL)
		return strReturn;

	m_fieldCtl = pRecordSet->Fields ->GetItem(long(iFieldIndex));
	varValue = m_fieldCtl->Value;
	if (varValue.vt == VT_NULL)
	{
		bstrValue = _T("");
	}
	else
	{
		bstrValue = varValue;
	}
	const char* pChar = bstrValue;
	strReturn = pChar;
	strReturn.TrimRight();
	return strReturn;
}

CString CAxManualDlg::GetFieldValue(_RecordsetPtr pRecordSet, CString strFieldName)
{
	CString strReturn = _T("");

	FieldPtr m_fieldCtl;
	int fieldCount = pRecordSet->Fields->Count;
	int fieldLength = 0;

	int nItem = 0;
	_variant_t varValue;
	_bstr_t bstrValue;

	_variant_t varName;
	_bstr_t bstrName;

	if( pRecordSet==NULL)
		return strReturn;

	for(int i=0; i<fieldCount;i++)
	{
		m_fieldCtl = pRecordSet->Fields ->GetItem(long(i));
		varValue = m_fieldCtl->Value;
		varName = m_fieldCtl->Name;

		if (varName.vt == VT_NULL)
		{
			bstrName = _T("");
		}
		else
		{
			bstrName = varName;
		}
		const char* pCharName = bstrName;
		CString strTemp = CString(pCharName);
		if(strFieldName.CompareNoCase(strTemp)==0)
		{
			if (varValue.vt == VT_NULL)
			{
				bstrValue = _T("");
			}
			else
			{
				bstrValue = varValue;
			}
			const char* pChar = bstrValue;
			strReturn = pChar;
			strReturn.TrimRight();
			break;
		}

	}
	return strReturn;
}

void CAxManualDlg::AdjustColumnWidth(CColorListCtrl* pList)
{
	pList->SetFont(&m_UserFont);
	pList->SetRedraw(FALSE);

	CHeaderCtrl* pHeaderCtrl = pList->GetHeaderCtrl();
	if(pHeaderCtrl!=NULL)
	{
		int nColumnCount = 	pHeaderCtrl->GetItemCount();

		for (int i = 0; i < nColumnCount; i++)
		{
			pList->SetColumnWidth(i, LVSCW_AUTOSIZE);
			int nColumnWidth = pList->GetColumnWidth(i);
			pList->SetColumnWidth(i, LVSCW_AUTOSIZE_USEHEADER);
			int nHeaderWidth = pList->GetColumnWidth(i); 
			pList->SetColumnWidth(i, max(nColumnWidth, nHeaderWidth));
		}
		pList->SetRedraw(TRUE);
	}
} 

BOOL CAxManualDlg::NetIsBroken()
{
	char name[255];
	gethostname(name, 255);
	struct hostent *remoteHost; 
	remoteHost = gethostbyname(name);
	BYTE* p = (BYTE*)remoteHost->h_addr;
	CString IP;
	IP.Format(_T("%d.%d.%d.%d"), p[0], p[1],p[2],p[3]);
	if(IP == _T("127.0.0.1"))
		return TRUE;

    return FALSE;
}

void CAxManualDlg::OnNMCustomDrawList(NMHDR *pNMHDR, LRESULT *pResult)
{
	UINT_PTR idfrom = pNMHDR->idFrom;
	int iCurMatch = idfrom - 1002;

	//类型安全转换   
	NMLVCUSTOMDRAW* pLVCD = reinterpret_cast<NMLVCUSTOMDRAW*>(pNMHDR);   
	*pResult = 0;   

	//指定列表项绘制前后发送消息   
	if(CDDS_PREPAINT == pLVCD->nmcd.dwDrawStage)   
	{   
		*pResult = CDRF_NOTIFYITEMDRAW;   
	}   
	else if(CDDS_ITEMPREPAINT == pLVCD->nmcd.dwDrawStage)   
	{   
		/*
		//奇数行   
		if(pLVCD->nmcd.dwItemSpec % 2)   
			pLVCD->clrTextBk = RGB(255, 255, 128);   
		//偶数行   
		else  
			pLVCD->clrTextBk = RGB(128, 255, 255);   

			*/
		//继续   
		*pResult = CDRF_DODEFAULT;   
	}   



	//if(idfrom == m_curMatchIndex+1002)
	//{
	//	//	LPNMCUSTOMDRAW pNMCD = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);
	//	NMLVCUSTOMDRAW* pLVCD = reinterpret_cast<NMLVCUSTOMDRAW*>(pNMHDR);

	////	*pResult = CDRF_DODEFAULT;

	//	if ( (CDDS_ITEMPREPAINT | CDDS_SUBITEM) == pLVCD->nmcd.dwDrawStage )
	//	{

	//		int iSel = m_ctrlPlayer[m_curMatchIndex].GetSelectionMark();

	//		POSITION pos = m_ctrlPlayer[m_curMatchIndex].GetFirstSelectedItemPosition();
	//		int nItem = static_cast<int>( pLVCD->nmcd.dwItemSpec );
	//		int index = m_ctrlPlayer[m_curMatchIndex].GetNextSelectedItem(pos);
	//		if (index == nItem)
	//		{
	//			COLORREF clrNewTextColor, clrNewBkColor;
	//			clrNewTextColor = RGB(255,255,255);  
	//			clrNewBkColor = RGB(49,106,197); 
	//			pLVCD->clrText = clrNewTextColor;
	//			pLVCD->clrTextBk = clrNewBkColor;
	//		}
	//	}
	//}
}

void CAxManualDlg::OnLvnItemchangedListPlayer(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
}

void CAxManualDlg::OnLvnItemchangedListEvent(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
}

void CAxManualDlg::OnBnClickedButtonRefresh()
{

  //  to_Upper(_T("啊啊12"));

	if(DataM()==NULL)
	{
		return;
	}

	CString strTable = _T("GA_Config");
	if(DataM()->TableIsExist(strTable))
	{
		CString strLanguageCode = DataM()->GetValue(strTable, _T("P_Lang"), 0);
		m_strLanguageCode = strLanguageCode;

	}

	m_dbString = getDBConnection();
	_ConnectionPtr pConnection;
	if(ConnectDB(m_dbString, pConnection))
	{

		AfxMessageBox(_T("数据库连接成功"));
	}
	else
	{
		AfxMessageBox(_T("数据库连接失败"));
		return;
	}

	CComboBox* pCombo = (CComboBox*)GetDlgItem(IDC_COMBO_STAGE);
	pCombo->ResetContent();
	
	CString strSQLString;
	strSQLString.Format(_T("%s"), m_sql_get_date);

	if(strSQLString == _T("")) return;

	_RecordsetPtr pRecordSet;
	if(!ExecSQL(strSQLString, pRecordSet))
		return;


	m_dateKey.clear();
	int iRow = 0;
	while(!pRecordSet->adoEOF)
	{
		CString strDate = GetFieldValue(pRecordSet, _T("Date"));
		CString strDateID = GetFieldValue(pRecordSet, _T("F_DisciplineDateID"));
		int idateid = _ttoi(strDateID);

		pCombo->AddString(strDate);

		m_dateKey.insert(make_pair(iRow, idateid));

		pRecordSet->MoveNext();
		iRow++;
	}

	UpdateData(FALSE);}

void CAxManualDlg::OnBnClickedButtonTriggerTeamPreview11()
{
	triggerStartList(0, _T("出发名单1"));
}

void CAxManualDlg::OnBnClickedButtonTriggerTeamPreview12()
{
	triggerStartList(0, _T("出发名单2"));
}

void CAxManualDlg::OnBnClickedButtonTriggerTeamPreview21()
{
	triggerStartList(1, _T("出发名单1"));
}

void CAxManualDlg::OnBnClickedButtonTriggerTeamPreview22()
{
	triggerStartList(1, _T("出发名单2"));
}
void CAxManualDlg::OnBnClickedButtonTriggerTeamPreview31()
{
	triggerStartList(2, _T("出发名单1"));
}

void CAxManualDlg::OnBnClickedButtonTriggerTeamPreview32()
{
	triggerStartList(2, _T("出发名单2"));
}
void CAxManualDlg::OnBnClickedButtonTriggerTeamPreview41()
{
	triggerStartList(3, _T("出发名单1"));
}

void CAxManualDlg::OnBnClickedButtonTriggerTeamPreview42()
{
	triggerStartList(3, _T("出发名单2"));
}
void CAxManualDlg::OnBnClickedButtonTriggerTeamPreview51()
{
	triggerStartList(4, _T("出发名单1"));
}

void CAxManualDlg::OnBnClickedButtonTriggerTeamPreview52()
{
	triggerStartList(4, _T("出发名单2"));
}
void CAxManualDlg::OnBnClickedButtonTriggerTeamPreview61()
{
	triggerStartList(5, _T("出发名单1"));
}

void CAxManualDlg::OnBnClickedButtonTriggerTeamPreview62()
{
	triggerStartList(5, _T("出发名单2"));
}
void CAxManualDlg::triggerStartList(int iCurMatch, CString cmdName)
{
	if(iCurMatch < m_MatchKey.size())
	{

		m_curMatchIndex = iCurMatch;
		m_MatchID = m_MatchKey[m_curMatchIndex];

		for(int i=0;i<m_MatchKey.size();i++)
		{
			int iRows = m_ctrlPlayer[i].GetItemCount();
			for(int j=0;j<iRows;j++)
			{
				m_ctrlPlayer[i].SetItemColor(j, RGB(0,0,0),RGB(184,184,184));
			}
		}

		m_ctrlPlayer[m_curMatchIndex].SetItemState(m_curPlayerIndex, LVIS_SELECTED|LVIS_FOCUSED, LVIS_SELECTED|LVIS_FOCUSED);
		m_ctrlPlayer[m_curMatchIndex].SetSelectionMark(m_curPlayerIndex);
		m_ctrlPlayer[m_curMatchIndex].SetItemColor(m_curPlayerIndex,RGB(255,255,255),RGB(10, 36, 106));

		int iSel = m_ctrlPlayer[m_curMatchIndex].GetSelectionMark();
		m_curPlayerIndex = iSel;
		m_PlayerBib = m_ctrlPlayer[m_curMatchIndex].GetItemText(iSel, 0);

		Param2DataCenter();

		// trigger preview cmd
		CString strCmd = GetTriggerCmd(cmdName);
		CmdTrigger()->TriggerCmd(strCmd, IAsCmdTrigger::emTriggerPreview);
	}
}

void CAxManualDlg::OnBnClickedButtonGetPlayerScore()
{
	resetMatchList();

	if(m_curMatchIndex<0)
	{
		return;
	}

	int iSel = m_ctrlPlayer[m_curMatchIndex].GetSelectionMark();
	if(iSel<0)
	{
		return;
	}

	AutoTriggerCmd();

}

void CAxManualDlg::OnBnClickedCheckAd()
{

	if(((CButton*)GetDlgItem(IDC_CHECK_AD))->GetCheck())
	{
		GetDlgItem(IDC_BUTTON_GET_TEAM_BUILD)->ShowWindow(FALSE);
		GetDlgItem(IDC_BUTTON_GET_ALLROUND_BUILD)->ShowWindow(FALSE);

		GetDlgItem(IDC_LIST_ALLROUD)->ShowWindow(TRUE);
		GetEventMedal();
	}
	else
	{
		if(m_stage == 1)
		{
			GetDlgItem(IDC_BUTTON_GET_TEAM_BUILD)->ShowWindow(TRUE);
			GetDlgItem(IDC_BUTTON_GET_ALLROUND_BUILD)->ShowWindow(TRUE);
		}
		if(m_stage == 4)
		{
			GetDlgItem(IDC_BUTTON_GET_TEAM_BUILD)->ShowWindow(TRUE);
		}

		GetDlgItem(IDC_LIST_ALLROUD)->ShowWindow(FALSE);
	}

	m_skin.SubClassWindowChildItem(m_hWnd);

}



void CAxManualDlg::OnBnClickedButtonGetTeamBuild()
{
	CString strCmd = GetTriggerCmd(_T("团体DE"));
	if(GetTeamScore()>1)
	{
		strCmd = GetTriggerCmd(_T("团体Build"));
	}
	CmdTrigger()->TriggerCmd(strCmd, IAsCmdTrigger::emTriggerPreview);

}

void CAxManualDlg::OnBnClickedButtonGetAllroundBuild()
{
	CString strCmd = GetTriggerCmd(_T("全能DE"));
	if(GetPlayerAllRoundScore()>1)
	{
		CString strCmd = GetTriggerCmd(_T("全能Build"));
	}
	CmdTrigger()->TriggerCmd(strCmd, IAsCmdTrigger::emTriggerPreview);
}
