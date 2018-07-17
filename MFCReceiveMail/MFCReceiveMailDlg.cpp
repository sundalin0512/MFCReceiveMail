
// MFCReceiveMailDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "MFCReceiveMail.h"
#include "MFCReceiveMailDlg.h"
#include "afxdialogex.h"
#include "ProgressBar.h"
#include "tool.h"
#include <fstream>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

CMFCReceiveMailDlg* g_pDlg = NULL;

// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
	public:
	CAboutDlg();

	// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
	protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CMFCReceiveMailDlg �Ի���



CMFCReceiveMailDlg::CMFCReceiveMailDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CMFCReceiveMailDlg::IDD, pParent)
	, m_cstrAccount(_T(""))
	, m_cstrPassword(_T(""))
	, m_cstrSeverName(_T(""))
	, m_cstrAddName(_T(""))
	, m_cstrDownloadDir(_T(""))
	//, m_cstrWhichClass(_T(""))
	, m_STime(COleDateTime::GetCurrentTime())
	, m_ETime(COleDateTime::GetCurrentTime())
	, m_cstrDownloadStatus(_T(""))
	, m_cstrClass(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_cstrAccount = _T("asm51asm@163.com");
	m_cstrPassword = _T("xxxxxxxxxxxxxxxxx");
	m_cstrSeverName = _T("pop.163.com");
	
	m_nRadioBn = 0;

	m_nStuCountPos = -1;
	m_nSureStuCountPos = -1;
	m_nSureStuNamePos = -1;
	m_nNotSureStuNamePos = -1;

	m_nSureStuCount = 0;
	m_nNotSureStuCount = 0;

	m_pDlgBar = NULL;
	m_bInquiryFlag = FALSE;
	m_bOnlySaveOnce = FALSE;
	m_bResaveFlag = FALSE;


	std::fstream stream("data.bin", std::ios_base::in | std::ios_base::out | std::ios_base::binary);
	
	size_t size;
	stream.read((char*)&size, sizeof(size_t));
	wchar_t *strPath = new wchar_t[size+1];
	stream.read((char*)strPath, size*2);
	strPath[size] = '\0';
	m_cstrDownloadDir.SetString(strPath);
	stream.close();
}

void CMFCReceiveMailDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

	DDX_Text(pDX, IDC_EDIT_Account, m_cstrAccount);
	DDX_Text(pDX, IDC_EDIT_Pwd, m_cstrPassword);
	DDX_Text(pDX, IDC_EDIT_Server, m_cstrSeverName);
	DDX_Text(pDX, IDC_EDIT_AddStu, m_cstrAddName);
	DDX_Text(pDX, IDC_EDIT_DownloadAddr, m_cstrDownloadDir);
	DDX_Control(pDX, IDC_LIST_ReportInfo, m_listctrShowMail);
	DDX_Control(pDX, IDC_DATETIMEPICKER_StartDate, m_dateCtrlStart);
	DDX_Control(pDX, IDC_DATETIMEPICKER_StartDate2, m_dateCtrlStartTime);
	DDX_Control(pDX, IDC_DATETIMEPICKER_EndDate, m_dateCtrlEnd);
	DDX_Control(pDX, IDC_DATETIMEPICKER_EndDate2, m_dateCtrlEndTime);
	DDX_Text(pDX, IDC_EDIT_Status, m_cstrDownloadStatus);
	DDX_Text(pDX, IDC_EDIT_Class, m_cstrClass);
	DDX_Control(pDX, IDC_EDIT_Status, m_ctlDownloadStatus);
	DDX_Control(pDX, IDC_Bn_Download, m_btnDownloadMail);
	DDX_Control(pDX, IDC_Bn_Inquiry, m_btnInquiry);
}

BEGIN_MESSAGE_MAP(CMFCReceiveMailDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_Bn_Inquiry, &CMFCReceiveMailDlg::OnBnClickedBnInquiry)
	ON_BN_CLICKED(IDC_Bn_Download, &CMFCReceiveMailDlg::OnBnClickedBnDownload)
	ON_BN_CLICKED(IDC_Bn_AddStu, &CMFCReceiveMailDlg::OnBnClickedBnAddstu)
	ON_NOTIFY(NM_RCLICK, IDC_LIST_ReportInfo, &CMFCReceiveMailDlg::OnNMRClickListReportinfo)
	ON_COMMAND(MENU_Operate_DelStu, &CMFCReceiveMailDlg::OnOperateDelstu)
	ON_MESSAGE(WM_Main_MSG, &CMFCReceiveMailDlg::OnMainMsg)
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_DATETIMEPICKER_StartDate2, &CMFCReceiveMailDlg::OnDtnDatetimechangeDatetimepickerStartdate2)
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_DATETIMEPICKER_EndDate2, &CMFCReceiveMailDlg::OnDtnDatetimechangeDatetimepickerEnddate2)
END_MESSAGE_MAP()


// CMFCReceiveMailDlg ��Ϣ�������

BOOL CMFCReceiveMailDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO:  �ڴ���Ӷ���ĳ�ʼ������

	//ע��jmail.dll
	TCHAR szPath[MAX_PATH] = { 0 };
	GetCurrentDirectory(MAX_PATH, szPath);
	lstrcat(szPath, _T("\\jmail.dll"));
	RegistryDll(CString(szPath));


	ConnectDB();
	InitListCtr();
	COleDateTime currentTime = COleDateTime::GetCurrentTime();
	COleDateTime preTime = currentTime - COleDateTimeSpan(1);
	if (currentTime.GetDayOfWeek() == 2)
		preTime = preTime - COleDateTimeSpan(2);
	CTime timeCurrent = CTime(currentTime.GetYear(), currentTime.GetMonth(), currentTime.GetDay(), currentTime.GetHour(), currentTime.GetMinute(), currentTime.GetMinute());
	CTime timeMiddle = CTime(currentTime.GetYear(), currentTime.GetMonth(), currentTime.GetDay(), 13, 0, 0);
	CTime timeBegin = CTime(preTime.GetYear(), preTime.GetMonth(), preTime.GetDay(), 0, 0, 0);
	CTime timeEnd = CTime(currentTime.GetYear(), currentTime.GetMonth(), currentTime.GetDay(), 0, 0, 0);
	if(timeCurrent > timeMiddle)
	{
		timeBegin = CTime(preTime.GetYear(), preTime.GetMonth(), preTime.GetDay(), 18, 0, 0);
		timeEnd = CTime(currentTime.GetYear(), currentTime.GetMonth(), currentTime.GetDay(), 13, 0, 0);
	}
	else
	{
		timeBegin = CTime(preTime.GetYear(), preTime.GetMonth(), preTime.GetDay(), 11, 0, 0);
		timeEnd = CTime(currentTime.GetYear(), currentTime.GetMonth(), currentTime.GetDay(), 1, 0, 0);
	}
	m_dateCtrlStart.SetTime(&timeBegin);
	m_dateCtrlStartTime.SetTime(&timeBegin);
	m_dateCtrlEnd.SetTime(&timeEnd);
	m_dateCtrlEndTime.SetTime(&timeEnd);

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CMFCReceiveMailDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CMFCReceiveMailDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

BOOL CMFCReceiveMailDlg::ConnectDB()
{
	int nRet = sqlite3_open(DB_NAME, &m_DB);

	if (nRet)
	{
		MessageBox(TEXT("Fail to connect database."));
	}
	else
	{
		OutputDebugString(TEXT("Successfully connect to the database."));
	}

	return TRUE;
}

BOOL CMFCReceiveMailDlg::InitListCtr()
{
	int nCol = 0;

	m_listctrShowMail.InsertColumn(nCol++, _T("ID"), 0, 50);
	m_listctrShowMail.InsertColumn(nCol++, _T("����"), 0, 60);
	m_listctrShowMail.InsertColumn(nCol++, _T("�༶"), 0, 60);
	m_listctrShowMail.InsertColumn(nCol++, _T("��������"), 0, 190);
	m_listctrShowMail.InsertColumn(nCol++, _T("·��"), 0, 200);
	m_listctrShowMail.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

	return TRUE;
}


BOOL CMFCReceiveMailDlg::AddEmployee()
{
	// ����Ƿ��������
	if (m_cstrClass.IsEmpty() && m_cstrAddName.IsEmpty())
	{
		MessageBox(_T("��ָ���༶������"));
		return FALSE;
	}

	if (m_cstrClass.IsEmpty())
	{
		MessageBox(_T("��ָ���༶"));
		return FALSE;
	}

	if (m_cstrAddName.IsEmpty())
	{
		MessageBox(_T("��ָ������"));
		return FALSE;
	}

	CString cstrID;
	CString cstrSql;

	// ͳһת�ɴ�д
	m_cstrClass.MakeUpper();

	cstrSql.Format(_T("INSERT INTO tb_Employee (employee_id, employee_name,\
						 employee_class) VALUES (NULL, \'%s\', \'%s\')"),
		m_cstrAddName, m_cstrClass);

	int nRet = SQLITE_ERROR; // �������շ���ֵ
	char *chErrMsg = NULL;   // �������մ��󷵻�ֵ
	char* bstrSql = Unicode2UTF8(cstrSql);
	nRet = sqlite3_exec(m_DB, bstrSql, NULL, NULL, &chErrMsg);
	if (nRet)
	{
		_bstr_t bstrError = chErrMsg;
		TRACE(bstrError.GetBSTR());
		return FALSE;
	}

	Inquiry();

	MessageBox(_T("����ɹ�\r\n"));
	return TRUE;
}


BOOL CMFCReceiveMailDlg::DownloadMail()
{
	CoInitialize(NULL);

	ReadDeadlineIni();


	try
	{
		jmail::IPOP3Ptr pPOP3("JMail.POP3");
		// TODO: ԭ����30s
		pPOP3->Timeout = 300; // �������ӷ�������ʱ���� 30S

		_bstr_t bstrAaccount = m_cstrAccount.GetBuffer();
		_bstr_t bstrPassword = m_cstrPassword.GetBuffer();
		_bstr_t bstrServerName = m_cstrSeverName.GetBuffer();
		pPOP3->Connect(bstrAaccount, bstrPassword, bstrServerName, 110); // �����ʼ���������110Ϊpop3Ĭ�϶˿ں�
		// ����ָ��
		jmail::IMessagesPtr pMessagesBox;
		// ��ȡ����
		pMessagesBox = pPOP3->Messages;

		long ncMail = NULL; // �ʼ���Ŀ,
		pMessagesBox->get_Count(&ncMail);

		ncMail--;
		if (ncMail <= 0)
		{
			throw - 1;
		}

		// ���ڿ�ʼ��һ���ûһ���ʼ�
		jmail::IMessagePtr pMail;
		//pMail->Charset = "UTF-8";
		int nMailCount = 0;
		for (int i = ncMail; i >= 1; --i)
		{ // �ʼ��±��1��ʼ
			TRACE("%d\r\n", i);

			SetDlgItemInt(IDC_EDIT_Status, (UINT)nMailCount);
			nMailCount++;

			try
			{
				pMail = pMessagesBox->Item[i];

				// 				pMail->Charset = "GB2312";
				// 				pMail->EnableCharsetTranslation = true;
				// 				pMail->ContentTransferEncoding = "base64";
				// 				pMail->Encoding = "base64";
				// 				pMail->ISOEncodeHeaders = true;
				// 				pMail->ContentType = "text/plain";
			}
			catch (_com_error e)
			{
				i = i + 1; // ���¶�ȡ
				continue;
			}

			DATE SendTime; // ���ڼ�¼�ʼ��ķ���ʱ��
			SendTime = pMail->GetDate();
			COleDateTime ColSendTime(SendTime); // ���ڼ�¼�ʼ��ķ���ʱ��
			_bstr_t bstrSendTime;
			bstrSendTime = ColSendTime.Format(_T("%Y-%m-%d %H:%M:%S"));

			COleDateTime ColRecvTime(SendTime); // ���ڼ�¼�ʼ��Ľ��ܵ�ʱ��
			_bstr_t bstrRecvTime;
			bstrRecvTime = ColSendTime.Format(_T("%Y-%m-%d %H:%M:%S"));

			// ���ʼ��ķ���ʱ�������ʼ���ʼ���յ�ʱ��ʱ���˳�����ֹ���������ʼ���
			COleDateTime colDateStart;
			COleDateTime colDateEnd;
			DownloadTime(colDateStart, colDateEnd);
			if (ColSendTime <= colDateStart)
			{
				break;
			}

			

			jmail::IAttachmentsPtr pAttchmentBox;
			pAttchmentBox = pMail->Attachments;
			long lcAttachment = NULL; // ���ڼ���ʼ��еĸ�����Ŀ
			pAttchmentBox->get_Count(&lcAttachment);

			// ���������ʼ��ĸ���
			for (int j = 0; j < lcAttachment; ++j)
			{
				jmail::IAttachmentPtr pAttchment;
				try
				{
					pAttchmentBox->get_Item(j, &pAttchment);// ��ȡ�ʼ�����
				}
				catch (_com_error e)
				{
					j = j - 1; // ���¶�ȡ
					continue;
				}

				_bstr_t bstrAttName = pAttchment->GetName();

				// �ж��Ƿ�ΪUTF8���룬����ǣ���UTF-8ת��ΪGB2312
				std::string strToConvert = bstrAttName;
				if (beUtf8((strToConvert.c_str())))
				{
					strToConvert = U2G(strToConvert.c_str());
					bstrAttName = strToConvert.c_str();
				}

				std::vector<std::string> TagNamePices;

				std::string strAttName = bstrAttName;
				std::string strSubject = strAttName; // ������

				// �зָ�������strSubject���ж������ĸ��༶
				std::vector<std::string> vecSegSubject;
				boost::split(vecSegSubject, strSubject, boost::is_any_of("_."));

				// �ж�strSubject��ǰ������ĸ�ǲ���CR
				// ���ʼ�����ʱ���Ƿ�����ѡ��Χ��
				BOOL BRet = IsCR(vecSegSubject, ColSendTime);
				CString cstrPath;

				if (BRet)
				{
					// ��ʼ�����ʼ������༶��������Ӧ�ļ���
					CheckDateCreateDir(vecSegSubject, cstrPath, ColSendTime);
				}
				else // �ʼ�����������
				{
					break;
				}


				// ��¼�������ص�ʱ��
				if (!m_bOnlySaveOnce)
				{
					DATE SendTimeLog; // ���ڼ�¼�ʼ��ķ���ʱ��
					SendTimeLog = pMail->GetDate();
					COleDateTime ColSendTimeLog(SendTimeLog); // ���ڼ�¼�ʼ��ķ���ʱ��
					m_bstrSendTimeLog = ColSendTimeLog.Format(_T("%Y-%m-%d %H:%M:%S"));

					m_bOnlySaveOnce = TRUE;
				}

				// �зָ�������strAttName���ж������ĸ��༶
				std::vector<std::string> vecSegTag;
				// �����ӷ��ָ�
				boost::split(vecSegTag, strAttName, boost::is_any_of("_."));

				_bstr_t	bstrOwner = vecSegTag[1].c_str();
				_bstr_t bstrPath = cstrPath;
				cstrPath += _T("\\");
				cstrPath += bstrAttName.GetBSTR();

				_bstr_t bstrFullName = cstrPath;

				try
				{
					pAttchment->SaveToFile(bstrFullName);
				}
				catch (_com_error e)
				{
					// ����ʧ�ܣ�˵������ͬ���ļ�
					// ��ȡ���Ѵ��ڵ��ļ���ʱ��
					MailIsRecv(bstrAttName);

					// m_bResaveFlagΪtrue�����������ʼ�
					if (m_bResaveFlag)
					{
						// ��ɾ��ԭ���ʼ��ĸ���
						BOOL bRet = DeleteFile(bstrFullName);
						if (!bRet)
						{
							// ɾ��ʧ��
							TRACE(_T("ɾ��ʧ��"));
						}

						pAttchment->SaveToFile(bstrFullName);
					}
					else
					{
						continue;
					}
				}

				// ����Ϣ�������ݿ�
				RecordMailInfo(bstrRecvTime, bstrAttName, bstrOwner, bstrPath);

				// ��ʾ�ʼ��������
				ShowMailDownloadStatus(bstrAttName, bstrOwner, bstrPath);

			}
			pAttchmentBox->Clear();
		}
		pMessagesBox->Clear();
		pPOP3->Disconnect();
	}
	catch (_com_error e)
	{
		TRACE(_T("ERROR"));
		OutputDebugString(e.Description());
	}
	catch (int e)
	{
		if (e > 0)
		{
			//MessageBox(TEXT("�������"));
		}
		else
		{
			TRACE(_T("���ճ����쳣"));
		}
	}
	CoUninitialize();

	return TRUE;
}

BOOL CMFCReceiveMailDlg::IsCR(std::vector<std::string> & strMailSubject, COleDateTime colMailSendTime)
{
	_bstr_t bstrClassName = strMailSubject[0].c_str();
	CString strClassName = bstrClassName;


	COleDateTime colDateStart;
	COleDateTime colDateEnd;

	DownloadTime(colDateStart, colDateEnd);

	CString cstrStartTime = colDateStart.Format(_T("%Y-%m-%d %H:%M:%S"));
	CString cstrEndTime = colDateEnd.Format(_T("%Y-%m-%d %H:%M:%S"));

	if (strClassName[0] == 'C' && strClassName[1] == 'R')
	{
		// ���ŶԱ�ʱ��
		if (colMailSendTime >= colDateStart && colMailSendTime <= colDateEnd)
		{
			return TRUE;
		}
	}
	else
	{
		return FALSE;
	}

	return FALSE;

}

BOOL CMFCReceiveMailDlg::DownloadTime(COleDateTime& colDateStart, COleDateTime& colDateEnd)
{
	int nYear = 0;
	int nMonth = 0;
	int nDay = 0;

	m_dateCtrlStart.GetTime(colDateStart);
	m_dateCtrlEnd.GetTime(colDateEnd);

	//if (m_strDeadline == _T("-1"))
	//{
	//	nYear = colDateStart.GetYear();
	//	nMonth = colDateStart.GetMonth();
	//	nDay = colDateStart.GetDay();

	//	colDateStart.SetDateTime(nYear, nMonth, nDay, 13, 0, 1);
	//}
	//else if (m_coleDeadlline > colDateStart && m_coleDeadlline < colDateEnd)
	//{
	//	colDateStart = m_coleDeadlline;
	//}
	//else if (m_coleDeadlline > colDateStart && m_coleDeadlline > colDateEnd)
	//{
	//	colDateStart = m_coleDeadlline;
	//}

	//nYear = colDateEnd.GetYear();
	//nMonth = colDateEnd.GetMonth();
	//nDay = colDateEnd.GetDay();
	//int nHour = colDateEnd.GetHour();
	//int nMin = colDateEnd.GetMinute();
	//int nSec = colDateEnd.GetSecond();

	return TRUE;
}


BOOL CMFCReceiveMailDlg::ReportTime(COleDateTime& colDateStart, COleDateTime& colDateEnd)
{
	int nYear = 0;
	int nMonth = 0;
	int nDay = 0;

	m_dateCtrlStart.GetTime(colDateStart);

	nYear = colDateStart.GetYear();
	nMonth = colDateStart.GetMonth();
	nDay = colDateStart.GetDay();
	colDateStart.SetDateTime(nYear, nMonth, nDay, 13, 0, 1);

	m_dateCtrlEnd.GetTime(colDateEnd);

	return TRUE;
}



BOOL CMFCReceiveMailDlg::CheckDateCreateDir(std::vector<std::string>  & strMailSubject, CString & strDir, COleDateTime ColSendTime)
{
	_bstr_t bstrClassName = strMailSubject[0].c_str();

	COleDateTime colDateEnd;
	m_dateCtrlEnd.GetTime(colDateEnd);
	int nYear = colDateEnd.GetYear();
	int nMonth = colDateEnd.GetMonth();
	int nDay = colDateEnd.GetDay();

	_bstr_t bstrHomeworkYear = nYear;
	CString strHomeworkMonth;
	strHomeworkMonth.Format(_T("%02d"), nMonth);
	CString strHomeworkDay;
	strHomeworkDay.Format(_T("%02d"), nDay);


	CreateDirectory(m_cstrDownloadDir, NULL);
	strDir += m_cstrDownloadDir;
	// 	namespace bf = boost::filesystem; // ����
	// 	bf::path path(strDir);
	// 	boost::filesystem::create_directories(path);
	strDir += _T("\\");
	strDir += (LPCSTR)bstrClassName;
	strDir += _T("\\");
	CreateDirectory(strDir, NULL);

	strDir += (LPCSTR)bstrClassName;
	strDir += _T("_");
	strDir += (LPCSTR)bstrHomeworkYear;
	strDir += _T("_");
	strDir += strHomeworkMonth;
	strDir += _T("_");
	strDir += strHomeworkDay;
	strDir += _T("\\");

	if (CreateDirectory(strDir, NULL))
	{
		// �����ɹ����ɹ���·����ӵ����ݿ���
		CString cstrSql;
		cstrSql.Format(_T("INSERT INTO tb_mailpath (mail_mailpathid, mail_path) VALUES (NULL, \'%s\')"), strDir);
		char* bstrSql = Unicode2UTF8(cstrSql);
		int rh = SQLITE_ERROR; // �������շ���ֵ
		char* cErrMsg = NULL; // �������մ���
		rh = sqlite3_exec(m_DB, bstrSql, NULL, NULL, &cErrMsg);
		if (rh)
		{
			_bstr_t bstrError = cErrMsg;
			TRACE(_T(":�����ʼ��ɹ���������Ϣ�޷��������ݿ�"));
			return FALSE;
		}
	}

	return TRUE;
}

static int MailIsRecvcallback(void *pWnd, int argc, char **argv, char **azColName)
{
	CMFCReceiveMailDlg* pDlg = (CMFCReceiveMailDlg*)pWnd;

	_bstr_t bstrTemp = argv[1];

	COleVariant coleTime(bstrTemp.GetBSTR());
	coleTime.ChangeType(VT_DATE);
	COleDateTime coleExistAttach = coleTime;

	COleDateTime colDateStart;
	COleDateTime colDateEnd;
	pDlg->DownloadTime(colDateStart, colDateEnd);

	if (coleExistAttach > colDateStart && coleExistAttach < colDateEnd)
	{
		pDlg->m_bResaveFlag = FALSE;
	}
	else
	{
		pDlg->m_bResaveFlag = TRUE;
	}

	return 0;
}

BOOL CMFCReceiveMailDlg::MailIsRecv(_bstr_t bstrFileName)
{
	CString cstrSql;
	cstrSql.Format(_T("SELECT mail_mailID, mail_sendtime, mail_name FROM tb_RecvMail WHERE mail_name = \'%s\'"), bstrFileName.GetBSTR());
	char* bstrSql = Unicode2UTF8(cstrSql);
	int rh = SQLITE_ERROR; // �������շ���ֵ
	BOOL bRet = FALSE;
	char* cErrMsg = NULL; // �������մ���
	rh = sqlite3_exec(m_DB, bstrSql, MailIsRecvcallback, this, &cErrMsg);
	if (rh)
	{
		_bstr_t bstrError = cErrMsg;
		TRACE(_T(":MailIsRecv"));
		return FALSE;
	}
	return bRet;
}

BOOL CMFCReceiveMailDlg::RecordMailInfo(_bstr_t RecvTime, _bstr_t strFileName, _bstr_t strOwnerName, _bstr_t strFilePath)
{
	// ����������
	CString cstrSql;
	CString temp;
	cstrSql.Format(_T("INSERT INTO tb_RecvMail (mail_mailID, mail_sendtime, mail_name, mail_SavePathID, mail_Ownerid) VALUES "));
	cstrSql += _T("(");
	temp.Format(_T("\'%s-%s\',"), RecvTime.GetBSTR(), strFileName.GetBSTR());
	cstrSql += temp;
	temp.Format(_T("\'%s\',"), RecvTime.GetBSTR());
	cstrSql += temp;
	temp.Format(_T("\'%s\',"), strFileName.GetBSTR());
	cstrSql += temp;
	temp.Format(_T("(SELECT mail_mailpathid FROM tb_mailpath WHERE mail_path=\'%s\'),"), strFilePath.GetBSTR());
	cstrSql += temp;
	temp.Format(_T(" (SELECT employee_id FROM tb_Employee WHERE employee_name=\'%s\')"), strOwnerName.GetBSTR());
	cstrSql += temp;
	cstrSql += _T(")");

	// ִ�в������
	char* bstrSql = Unicode2UTF8(cstrSql);
	int rh = SQLITE_ERROR; // �������շ���ֵ
	char* cErrMsg = NULL; // �������մ���
	//int nID; // ��������ȡԱ��ID

	rh = sqlite3_exec(m_DB, bstrSql, NULL, NULL, &cErrMsg);
	if (rh)
	{
		_bstr_t bstrError = cErrMsg;
		TRACE(_T(":�����ʼ��ɹ���������Ϣ�޷��������ݿ�"));
		return -1;
	}
	return FALSE;
}


static int FindClassCallback(void *Wnd, int argc, char **argv, char **azColName)
{
	CMFCReceiveMailDlg * pDlg = (CMFCReceiveMailDlg *)Wnd;
	_bstr_t bstrTemp = argv[2];

	pDlg->m_strExactlyClass = bstrTemp.GetBSTR();

	return 0;
}


static int ListMailInfocallback(void *Wnd, int argc, char **argv, char **azColName)
{
	CMFCReceiveMailDlg * pDlg = (CMFCReceiveMailDlg *)Wnd;

	pDlg->m_bInquiryFlag = TRUE;

	_bstr_t bstrTemp = argv[0];
	pDlg->m_listctrShowMail.InsertItem(0, bstrTemp);
	for (int i = 1; i <= argc; ++i)
	{
		if (i == 4)
		{
			bstrTemp = argv[i - 1];
		}
		else
		{
			bstrTemp = argv[i];
		}

		if (i == 2)
		{
			_bstr_t bstrId = argv[0];
			_bstr_t bstrName = argv[1];

			CString cstrSql;
			cstrSql.Format(_T("SELECT employee_id, employee_name, employee_class FROM \
							   tb_Employee WHERE employee_id = \'%s\' and employee_name = \'%s\'"),
				bstrId.GetBSTR(), bstrName.GetBSTR());

			char* bstrSql = Unicode2UTF8(cstrSql);
			char* cErrMsg = NULL; // �������մ���
			int rh = sqlite3_exec(pDlg->m_DB, bstrSql, FindClassCallback, pDlg, &cErrMsg);

			pDlg->m_listctrShowMail.SetItemText(0, i++, pDlg->m_strExactlyClass);

			if (bstrTemp.GetBSTR() == NULL)
			{
				break;
			}

			pDlg->m_listctrShowMail.SetItemText(0, i, bstrTemp);

			continue;
		}

		pDlg->m_listctrShowMail.SetItemText(0, i, bstrTemp);
	}
	return 0;
}


BOOL CMFCReceiveMailDlg::ListMailInfo(COleDateTime & STime, COleDateTime & ETime)
{
	try
	{
		_bstr_t bstrSTime = STime.Format(_T("%Y-%m-%d %H:%M:%S"));
		_bstr_t bstrETime = ETime.Format(_T("%Y-%m-%d %H:%M:%S"));

		CString cstrSql;
		cstrSql.Format(_T("SELECT employee_id, employee_name, mail_name, mail_path FROM tb_Employee LEFT OUTER JOIN (SELECT mail_Ownerid, mail_name, mail_path FROM  tb_RecvMail,\
						  tb_mailpath ON tb_RecvMail.mail_SavePathID = tb_mailpath.mail_mailpathid WHERE mail_sendtime >= \'%s\' and mail_sendtime <= \'%s\')ON employee_id = mail_Ownerid"),
			bstrSTime.GetBSTR(), bstrETime.GetBSTR());

		char* bstrSql = Unicode2UTF8(cstrSql);
		int rh = SQLITE_ERROR; // �������շ���ֵ
		char* cErrMsg = NULL; // �������մ���
		rh = sqlite3_exec(m_DB, bstrSql, ListMailInfocallback, this, &cErrMsg);
		if (rh)
		{
			_bstr_t bstrError = cErrMsg;
			TRACE(_T(":�����ʼ��ɹ���������Ϣ�޷��������ݿ�"));
			return -1;
		}
		return FALSE;
	}
	catch (_com_error &e)
	{
		TRACE(_T("ʱ��������Ϸ�"));
	}
	return FALSE;
}

BOOL CMFCReceiveMailDlg::SplitFileName(std::vector<std::string> & Tag_pices, std::string & strFileName)
{
	boost::split(Tag_pices, strFileName, boost::is_any_of("_."));

	int nId = -1;
	int nID = GetEmloyeeID(Tag_pices[1].c_str());
	if (nID < 0)
	{
		return FALSE;
	}

	return TRUE;
}

static int AddEmploycallback(void *NotUsed, int argc, char **argv, char **azColName)
{
	*(int *)NotUsed = boost::lexical_cast<int>(*argv);
	return 0;
}

int CMFCReceiveMailDlg::GetEmloyeeID(_bstr_t bstrEmpName)
{
	CString cstrSql;
	cstrSql.Format(_T("SELECT employee_id FROM tb_Employee WHERE employee_name=\'%s\'"), bstrEmpName.GetBSTR());

	char* bstrSql = Unicode2UTF8(cstrSql);
	int rh = SQLITE_ERROR; // �������շ���ֵ
	char* cErrMsg = NULL; // �������մ���
	int nID = -1; // ��������ȡԱ��ID
	rh = sqlite3_exec(m_DB, bstrSql, AddEmploycallback, &nID, &cErrMsg);
	if (rh)
	{
		_bstr_t bstrError = cErrMsg;
		TRACE(bstrError.GetBSTR());
		return -1;
	}

	return nID;
}

static int WriteFilecallback(void *hFile, int argc, char **argv, char **azColName)
{
	// argv[0]		ID
	// argv[1]		Name
	// argv[2]		Class
	// argv[3]		CR27_xxx_2017.10.07.rar
	// argv[4]		�ļ�Ŀ¼

	CString Str;
	_bstr_t bstrName = argv[1];
	_bstr_t bstrFile = argv[3];

	CString strName = bstrName;
	CString strFile = bstrFile;

	DWORD nBytes;
	static int nSureNameCount = 0;
	static int nNotSureNameCount = 0;
	// ���û����ҵ
	if (bstrFile.GetBSTR() == NULL)
	{
		g_pDlg->m_nNotSureStuCount++;

		DWORD dwRetPos = SetFilePointer(hFile, g_pDlg->m_nNotSureStuNamePos, NULL, FILE_BEGIN);
		int nPos = LOWORD(dwRetPos);
		WriteFile(hFile, strName.GetBuffer(), strName.GetLength() * sizeof(TCHAR), &nBytes, NULL);
		nNotSureNameCount++;
		g_pDlg->m_nNotSureStuNamePos += INTERVAL * INTERVAL_PLUS;

		if (nNotSureNameCount == ONE_LINE_NAMES)
		{
			CString tmpStr = _T("\r\n");
			WriteFile(hFile, tmpStr.GetBuffer(), tmpStr.GetLength() * sizeof(TCHAR), &nBytes, NULL);

			DWORD dwRetPos = SetFilePointer(hFile, 0, NULL, FILE_CURRENT);
			g_pDlg->m_nNotSureStuNamePos = LOWORD(dwRetPos);

			nNotSureNameCount = 0;
		}
	}

	// ���������ҵ
	if (bstrFile.GetBSTR() != NULL)
	{
		// �ж��ʼ�������������İ༶�Ƿ���ȷ
		std::vector<std::string> Tag_pices;
		std::string strAttachName = argv[3];
		boost::split(Tag_pices, strAttachName, boost::is_any_of("_."));
		_bstr_t bstrSplitAttachClass = Tag_pices[0].c_str();
		CString strSplitAttachClass = bstrSplitAttachClass.GetBSTR();
		_bstr_t bstrAttachCls = argv[2];
		CString strAttachCls = bstrAttachCls.GetBSTR();
		if (strSplitAttachClass == strAttachCls)
		{
			g_pDlg->m_nSureStuCount++;

			DWORD dwRetPos = SetFilePointer(hFile, g_pDlg->m_nSureStuNamePos, NULL, FILE_BEGIN);
			int nPos = LOWORD(dwRetPos);
			WriteFile(hFile, strName.GetBuffer(), strName.GetLength() * sizeof(TCHAR), &nBytes, NULL);
			nSureNameCount++;

			g_pDlg->m_nSureStuNamePos += INTERVAL * INTERVAL_PLUS;

			if (nSureNameCount == ONE_LINE_NAMES)
			{
				CString tmpStr = _T("\r\n");
				WriteFile(hFile, tmpStr.GetBuffer(), tmpStr.GetLength() * sizeof(TCHAR), &nBytes, NULL);

				DWORD dwRetPos = SetFilePointer(hFile, 0, NULL, FILE_CURRENT);
				g_pDlg->m_nSureStuNamePos = LOWORD(dwRetPos);

				nSureNameCount = 0;
			}
		}
	}

	return 0;
}

BOOL MatchCatalog(CMFCReceiveMailDlg* pDlg, _bstr_t bstrDir, /*output*/std::string& strPathClass)
{
	std::vector<std::string> Tag_pices;
	std::string strDir = bstrDir;

	boost::split(Tag_pices, strDir, boost::is_any_of("\\_"));

	std::string strDay = Tag_pices[Tag_pices.size() - 2];
	int nCatalogDay = boost::lexical_cast<int>(strDay);
	std::string strMonth = Tag_pices[Tag_pices.size() - 3];
	int nCatalogMonth = boost::lexical_cast<int>(strMonth);
	std::string strYear = Tag_pices[Tag_pices.size() - 4];
	int nCatalogYear = boost::lexical_cast<int>(strYear);
	std::string strClass = Tag_pices[Tag_pices.size() - 5];

	COleDateTime colDateStart;
	COleDateTime colDateEnd;
	pDlg->m_dateCtrlStart.GetTime(colDateStart);
	pDlg->m_dateCtrlEnd.GetTime(colDateEnd);

	int nNeedDay = colDateEnd.GetDay();
	int nNeedMonth = colDateEnd.GetMonth();
	int nNeedYear = colDateEnd.GetYear();

	strPathClass = strClass;

	if (nCatalogDay == nNeedDay &&
		nCatalogMonth == nNeedMonth &&
		nCatalogYear == nNeedYear)
	{
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

static int GenStatmentcallback(void *pWnd, int argc, char **argv, char **azColName)
{
	CMFCReceiveMailDlg* pDlg = (CMFCReceiveMailDlg *)pWnd;
	g_pDlg = pDlg;
	_bstr_t bstrDir = argv[0];

	// ƥ����Ҫ���ɱ����Ŀ¼
	std::string strPathClass;
	BOOL bMatchRet = MatchCatalog(pDlg, bstrDir, strPathClass);
	if (!bMatchRet)
	{
		return 0;
	}

	CString strSqlClass(strPathClass.c_str());

	_bstr_t bstrFilepath = bstrDir + _T("�ձ���.txt");

	HANDLE hFile;
	hFile = CreateFile(bstrFilepath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
	{
		return -1;
	}


	COleDateTime colDateStart;
	COleDateTime colDateEnd;
	pDlg->ReportTime(colDateStart, colDateEnd);

	// ���ѧ������
	pDlg->m_nSureStuCount = 0;
	pDlg->m_nNotSureStuCount = 0;


	// ��txt�ĵ���ͷд���ֽ��ļ�ͷ0xfeff��ֹ����
	DWORD nBytes;
	WORD a = 0xFEFF;
	WriteFile(hFile, &a, sizeof(a), &nBytes, NULL);

	SetFilePointer(hFile, CAPTION_POS, NULL, FILE_BEGIN);

	// д�����
	CString str = _T("ÿ��ѧԱ��ҵ�ύ��\r\n\r\n");
	WriteFile(hFile, str, str.GetLength() * sizeof(TCHAR), &nBytes, NULL);

	// д��༶
	CString strClass;
	strClass.Format(_T("�༶: %s"), strSqlClass);
	WriteFile(hFile, strClass, strClass.GetLength() * sizeof(TCHAR), &nBytes, NULL);

	COleDateTime CurTime = COleDateTime::GetCurrentTime();
	_bstr_t bstrCurTime = CurTime.Format(TEXT("%Y-%m-%d %H:%M:%S"));
	CString strCurTime;
	strCurTime.Format(_T("����:%s"), bstrCurTime.GetBSTR());

	// д������
	SetFilePointer(hFile, INTERVAL, NULL, FILE_CURRENT);
	WriteFile(hFile, strCurTime, strCurTime.GetLength() * sizeof(TCHAR), &nBytes, NULL);

	// ��¼д��ѧԱ������λ��
	DWORD dwRetPos = SetFilePointer(hFile, INTERVAL, NULL, FILE_CURRENT);
	pDlg->m_nStuCountPos = LOWORD(dwRetPos);

	// ռλ
	CString strSpace = _T("  ");
	for (int j = 0; j < MINI_HOW_MANG_WORDS; j++)
	{
		WriteFile(hFile, strSpace, strSpace.GetLength() * sizeof(TCHAR), &nBytes, NULL);
	}

	CString strNewLine = _T("\r\n\r\n");
	WriteFile(hFile, strNewLine, strNewLine.GetLength() * sizeof(TCHAR), &nBytes, NULL);

	// �ѽ�ѧԱ
	CString strSureStu = _T("�ѽ�ѧԱ:\r\n");
	WriteFile(hFile, strSureStu, strSureStu.GetLength() * sizeof(TCHAR), &nBytes, NULL);
	dwRetPos = SetFilePointer(hFile, INTERVAL, NULL, FILE_CURRENT);
	pDlg->m_nSureStuNamePos = LOWORD(dwRetPos);

	// ռλ
	strNewLine = _T("\r\n");
	for (int i = 0; i < LINES_COUNT; i++)
	{
		for (int j = 0; j < HOW_MANG_WORDS; j++)
		{
			WriteFile(hFile, strSpace, strSpace.GetLength() * sizeof(TCHAR), &nBytes, NULL);
		}

		WriteFile(hFile, strNewLine, strNewLine.GetLength() * sizeof(TCHAR), &nBytes, NULL);
	}

	// δ��ѧԱ
	CString strNotSureStu = _T("δ��ѧԱ:\r\n");
	WriteFile(hFile, strNotSureStu, strNotSureStu.GetLength() * sizeof(TCHAR), &nBytes, NULL);
	dwRetPos = SetFilePointer(hFile, INTERVAL, NULL, FILE_CURRENT);
	pDlg->m_nNotSureStuNamePos = LOWORD(dwRetPos);

	// ��ѯ���ݿ��ʱ�䷶Χ
	_bstr_t bstrSTime = colDateStart.Format(TEXT("%Y-%m-%d %H:%M:%S"));
	_bstr_t bstrETime = colDateEnd.Format(TEXT("%Y-%m-%d %H:%M:%S"));

	CString cstrSql1, cstrSql2;
	cstrSql1.Format(_T("SELECT employee_id, employee_name, employee_class, mail_name, mail_path FROM tb_Employee LEFT OUTER JOIN (SELECT mail_Ownerid, mail_name, mail_path FROM  tb_RecvMail,\
						tb_mailpath ON tb_RecvMail.mail_SavePathID = tb_mailpath.mail_mailpathid WHERE mail_sendtime >= \'%s\' and mail_sendtime <= \'%s\')ON employee_id = mail_Ownerid where employee_class = \'%s\'"),
		bstrSTime.GetBSTR(), bstrETime.GetBSTR(), strSqlClass);

	char* bstrSql = Unicode2UTF8(cstrSql1);
	int rh = SQLITE_ERROR; // �������շ���ֵ
	char* cErrMsg = NULL; // �������մ���
	rh = sqlite3_exec(pDlg->m_DB, bstrSql, WriteFilecallback, hFile, &cErrMsg);

	// д��ѧԱ����
	SetFilePointer(hFile, pDlg->m_nStuCountPos, NULL, FILE_BEGIN);
	CString strStuCount = _T("ѧԱ����:");
	WriteFile(hFile, strStuCount, strStuCount.GetLength() * sizeof(TCHAR), &nBytes, NULL);

	strStuCount.Format(_T("%d"), pDlg->m_nSureStuCount + pDlg->m_nNotSureStuCount);
	WriteFile(hFile, strStuCount, strStuCount.GetLength() * sizeof(TCHAR), &nBytes, NULL);

	// д���ѽ���ҵ��ѧԱ����
	SetFilePointer(hFile, INTERVAL_PLUS, NULL, FILE_CURRENT);
	CString strSureStuCount = _T("�ѽ���ҵ��ѧԱ����:");
	WriteFile(hFile, strSureStuCount, strSureStuCount.GetLength() * sizeof(TCHAR), &nBytes, NULL);
	strSureStuCount.Format(_T("%d"), pDlg->m_nSureStuCount);
	WriteFile(hFile, strSureStuCount, strSureStuCount.GetLength() * sizeof(TCHAR), &nBytes, NULL);

	CloseHandle(hFile);
	if (rh)
	{
		_bstr_t bstrError = cErrMsg;
		TRACE(_T("GenStatmentcallback ���ִ���"));
		return -1;
	}
	return FALSE;
}

int CMFCReceiveMailDlg::GenStatment()
{
	CString cstrSql;
	cstrSql.Format(_T("SELECT mail_path FROM tb_mailpath"));

	char* bstrSql = Unicode2UTF8(cstrSql);
	int rh = SQLITE_ERROR; // �������շ���ֵ
	char* cErrMsg = NULL; // �������մ���
	int nID = -1; // ��������ȡԱ��ID
	rh = sqlite3_exec(m_DB, bstrSql, GenStatmentcallback, this, &cErrMsg);
	if (rh)
	{
		_bstr_t bstrError = cErrMsg;
		TRACE(bstrError.GetBSTR());
		return -1;
	}
	return nID;
}

void CMFCReceiveMailDlg::Inquiry()
{
	m_listctrShowMail.DeleteAllItems();

	COleDateTime colDateStart;
	COleDateTime colDateEnd;
	m_dateCtrlStart.GetTime(colDateStart);
	m_dateCtrlEnd.GetTime(colDateEnd);

	int nYearS = colDateStart.GetYear();
	int nMonthS = colDateStart.GetMonth();
	int nDayS = colDateStart.GetDay();

	int nYearE = colDateEnd.GetYear();
	int nMonthE = colDateEnd.GetMonth();
	int nDayE = colDateEnd.GetDay();

	colDateStart.SetDateTime(nYearS, nMonthS, nDayS, 13, 0, 1);

	ListMailInfo(colDateStart, colDateEnd);
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CMFCReceiveMailDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CMFCReceiveMailDlg::OnBnClickedBnInquiry()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	UpdateData(TRUE);

	m_bInquiryFlag = FALSE;

	Inquiry();

	if (!m_bInquiryFlag)
	{
		AfxMessageBox(_T("δ�ҵ�������ݣ�����༶��ʱ����Ƿ���ȷ"));
	}
}


UINT AFX_CDECL DownloadMailPROC(LPVOID lparam)
{
	CMFCReceiveMailDlg *pDlg = ((CMFCReceiveMailDlg*)lparam);
	BOOL BRet = pDlg->DownloadMail();

	pDlg->m_ctlDownloadStatus.SetWindowTextW(_T("over......"));

	pDlg->m_pDlgBar->ShowWindow(SW_HIDE);

	if (BRet)
	{
		pDlg->MessageBox(TEXT("�������"));
		ShellExecute(NULL, _T("explore"), pDlg->m_cstrDownloadDir, NULL, NULL, SW_SHOW);

	}

	if (pDlg->m_bOnlySaveOnce)
	{
		// ���汾�������ʼ��Ľ�ֹ���ڣ��´����ؾʹ��������֮��ʼ����
		CString strBSTR;
		strBSTR.Format(_T("%s"), pDlg->m_bstrSendTimeLog.GetBSTR());

		CString strPath;
		DWORD dwReadBytes = GetCurrentDirectory(MAXBYTE, strPath.GetBufferSetLength(MAXBYTE));
		strPath.ReleaseBuffer(dwReadBytes);
		strPath += _T("\\DeadlineLog.ini");
		BOOL bRet = WritePrivateProfileString(_T("��ֹʱ��"),
			_T("_bstr_t"),
			strBSTR,
			strPath);

		if (!bRet)
		{
			TRACE(_T("�����ֹʱ��ʧ��"));
		}
	}

	// ���ɱ���
	pDlg->GenStatment();

	pDlg->Inquiry();

	pDlg->m_btnDownloadMail.EnableWindow(TRUE);
	pDlg->m_btnInquiry.EnableWindow(TRUE);

	return 0;
}


void CMFCReceiveMailDlg::OnBnClickedBnDownload()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	UpdateData(TRUE);
	std::fstream stream("data.bin", std::ios_base::in | std::ios_base::out | std::ios_base::trunc | std::ios_base::binary);
	if(stream.is_open())
	{
		wstring tmppath = m_cstrDownloadDir.GetString();
		size_t size = tmppath.size();
		stream.write((char*)&size, sizeof(size_t));
		stream.write((char*)tmppath.c_str(), 2 * tmppath.size());
		//stream << tmppath.c_str();
		stream.close();
	}
	// �������
	COleDateTime colDateTimeS;
	m_dateCtrlStart.GetTime(colDateTimeS);

	COleDateTime colDateTimeE;
	m_dateCtrlEnd.GetTime(colDateTimeE);

	if (colDateTimeS >= colDateTimeE)
	{
		AfxMessageBox(_T("ʱ�䷶Χ����ȷ"));
		m_btnDownloadMail.EnableWindow(TRUE);
		m_btnInquiry.EnableWindow(TRUE);
		return;
	}

	COleDateTime colDateStart;
	COleDateTime colDateEnd;
	m_dateCtrlStart.GetTime(colDateStart);
	//int nYear = colDateStart.GetYear();
	//int nMonth = colDateStart.GetMonth();
	//int nDay = colDateStart.GetDay();
	//int nHour = colDateStart.GetHour();
	//int nMinute = colDateStart.GetMinute();
	//int nSecond = colDateStart.GetSecond();

	//colDateStart.SetDateTime(nYear, nMonth, nDay, nHour, nMinute, nSecond);

	m_dateCtrlEnd.GetTime(colDateEnd);
	CString cstrStartTime = colDateStart.Format(_T("%Y-%m-%d %H:%M:%S"));
	CString cstrEndTime = colDateEnd.Format(_T("%Y-%m-%d %H:%M:%S"));

	CString strText;
	strText.Format(_T("������Ϣ:\r\n\tʱ��� : %s -- %s"),
		cstrStartTime,
		cstrEndTime);

	int nRet = ::MessageBox(NULL, strText, _T("��ȷ������������"), MB_YESNO);
	if (nRet == IDYES)
	{
		// ��ģ̬�Ի���
		m_pDlgBar = new CProgressBar;
		m_pDlgBar->Create(IDD_ProgressBar);
		m_pDlgBar->ShowWindow(SW_SHOW);

		m_btnDownloadMail.EnableWindow(FALSE);
		m_btnInquiry.EnableWindow(FALSE);

		m_listctrShowMail.DeleteAllItems();

		CWinThread* pWndThread = AfxBeginThread(DownloadMailPROC, this);
		m_hDownloadThread = pWndThread->m_hThread;

		m_ctlDownloadStatus.SetWindowTextW(_T("Downloading......"));
	}
}


void CMFCReceiveMailDlg::OnBnClickedBnAddstu()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	UpdateData(TRUE);
	AddEmployee();
}


void CMFCReceiveMailDlg::OnNMRClickListReportinfo(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	*pResult = 0;

	int nIndex = pNMItemActivate->iItem;
	if (nIndex == -1)
	{
		return;
	}

	CMenu mn;
	// ���ز˵�
	mn.LoadMenu(MENU_Operate);
	// ��ȡ�������Ӳ˵�
	CMenu* pMn = mn.GetSubMenu(0);

	CPoint pt;
	GetCursorPos(&pt);

	// ���ͻ�������ת��Ϊ��Ļ����
	//ClientToScreen();

	// �����˵�
	pMn->TrackPopupMenu(TPM_LEFTALIGN, pt.x, pt.y, this);
}


void CMFCReceiveMailDlg::OnOperateDelstu()
{
	// TODO:  �ڴ���������������
	int nCurSel = m_listctrShowMail.GetSelectionMark();

	CString strStuId = m_listctrShowMail.GetItemText(nCurSel, 0);
	CString strName = m_listctrShowMail.GetItemText(nCurSel, 1);
	CString strClass = m_listctrShowMail.GetItemText(nCurSel, 2);

	CString cstrSql;

	// ��Ҫɾ��ĳ��д����
	// ����ɾ����ѧ������ҵ�ύ��¼��
	// ��Ϊ��ҵ�ύ��¼�ı����汣���˸�ѧ����ѧ�ţ�ֱ��ɾ��ѧ���������������
	// ����������Ǳ�������ҵ�ύ��¼��ѧ����ѧ��
	cstrSql.Format(_T("DELETE FROM tb_RecvMail where mail_Ownerid = \'%s\'"),
		strStuId);
	char* bstrSql = Unicode2UTF8(cstrSql);
	int nRet = SQLITE_ERROR; // �������շ���ֵ
	char *chErrMsg = NULL;   // �������մ��󷵻�ֵ
	nRet = sqlite3_exec(m_DB, bstrSql, NULL, NULL, &chErrMsg);
	if (nRet)
	{
		_bstr_t bstrError = chErrMsg;
		TRACE(bstrError.GetBSTR());
		return;
	}
	else if (nRet == SQLITE_OK)
	{
		cstrSql.Format(_T("DELETE FROM tb_Employee where\
						  employee_class = \'%s\' and employee_name = \'%s\'"),
			strClass, strName);
		bstrSql = Unicode2UTF8(cstrSql);
		nRet = sqlite3_exec(m_DB, bstrSql, NULL, NULL, &chErrMsg);
		if (nRet)
		{
			_bstr_t bstrError = chErrMsg;
			TRACE(bstrError.GetBSTR());
			return;
		}

		Inquiry();

		MessageBox(_T("ɾ���ɹ�\r\n"));
	}

	return;
}


afx_msg LRESULT CMFCReceiveMailDlg::OnMainMsg(WPARAM wParam, LPARAM lParam)
{
	TerminateThread(m_hDownloadThread, 0);

	m_ctlDownloadStatus.SetWindowTextW(_T("over......"));
	m_btnDownloadMail.EnableWindow(TRUE);

	return 0;
}

BOOL CMFCReceiveMailDlg::ReadDeadlineIni()
{
	CString strPath;
	DWORD dwReadBytes = GetCurrentDirectory(MAXBYTE, strPath.GetBufferSetLength(MAXBYTE));
	strPath.ReleaseBuffer(dwReadBytes);
	strPath += _T("\\DeadlineLog.ini");
	dwReadBytes = GetPrivateProfileString(_T("��ֹʱ��"),
		_T("_bstr_t"),
		_T("-1"),
		m_strDeadline.GetBufferSetLength(MAXLEN_IFDESCR),
		MAXLEN_IFDESCR,
		strPath);
	m_strDeadline.ReleaseBuffer();

	if (m_strDeadline == _T("-1"))
	{
		return FALSE;
	}

	COleVariant coleTime(m_strDeadline);
	coleTime.ChangeType(VT_DATE);
	m_coleDeadlline = coleTime;

	return TRUE;
}

void CMFCReceiveMailDlg::ShowMailDownloadStatus(_bstr_t strFileName, _bstr_t strOwnerName, _bstr_t strFilePath)
{
	m_listctrShowMail.InsertItem(0, _T(""));
	m_listctrShowMail.SetItemText(0, 1, strOwnerName);
	m_listctrShowMail.SetItemText(0, 3, strFileName);
	m_listctrShowMail.SetItemText(0, 4, strFilePath);
}



void CMFCReceiveMailDlg::OnDtnDatetimechangeDatetimepickerStartdate2(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMDATETIMECHANGE pDTChange = reinterpret_cast<LPNMDATETIMECHANGE>(pNMHDR);
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	COleDateTime date;
	m_dateCtrlStart.GetTime(date);
	COleDateTime time;
	m_dateCtrlStartTime.GetTime(time);
	date.SetDateTime(date.GetYear(),date.GetMonth(),date.GetDay(),time.GetHour(), time.GetMinute(), time.GetSecond());
	m_dateCtrlStart.SetTime(date);
	*pResult = 0;
}


void CMFCReceiveMailDlg::OnDtnDatetimechangeDatetimepickerEnddate2(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMDATETIMECHANGE pDTChange = reinterpret_cast<LPNMDATETIMECHANGE>(pNMHDR);
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	COleDateTime date;
	m_dateCtrlEnd.GetTime(date);
	COleDateTime time;
	m_dateCtrlEndTime.GetTime(time);
	date.SetDateTime(date.GetYear(), date.GetMonth(), date.GetDay(), time.GetHour(), time.GetMinute(), time.GetSecond());
	m_dateCtrlEnd.SetTime(date);
	*pResult = 0;
}
