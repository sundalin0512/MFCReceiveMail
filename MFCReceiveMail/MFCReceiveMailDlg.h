
// MFCReceiveMailDlg.h : ͷ�ļ�
//

#pragma once

#define DB_NAME	"RecvMail"

#import "jmail.dll"
#include "sqlite3.h"
#include "boost/algorithm/string.hpp"
#include "boost/lexical_cast.hpp"
#include <iostream>
#include "vector"
#include "afxwin.h"
#include "ProgressBar.h"
#include "aes.h"

using namespace std;

#define		ONE_WORD_BYTE			4
#define		CAPTION_POS				((ONE_WORD_BYTE) * 15)
#define		INTERVAL				ONE_WORD_BYTE
#define		INTERVAL_PLUS			8
#define		ONE_LINE_NAMES			5
#define     LINES_COUNT				4
#define		HOW_MANG_WORDS			120
#define		MINI_HOW_MANG_WORDS		40

// CMFCReceiveMailDlg �Ի���
class CMFCReceiveMailDlg : public CDialogEx
{
// ����
public:
	CMFCReceiveMailDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_MFCRECEIVEMAIL_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	CString m_cstrAccount;
	CString m_cstrPassword;
	CString m_cstrSeverName;
	CString m_cstrAddName;
	CString m_cstrDownloadDir;
	// ���ء���ѯ�༶
	CString m_cstrDownloadStatus;
	CString m_cstrClass;// Ҫ���ѧԱ���ڵİ༶
	CString m_strDeadline;
	CString m_strExactlyClass;

	_bstr_t m_bstrSendTimeLog;

	COleDateTime m_coleDeadlline;
	COleDateTime m_STime;
	COleDateTime m_ETime;
	CDateTimeCtrl m_dateCtrlStart;
	CDateTimeCtrl m_dateCtrlStartTime;
	CDateTimeCtrl m_dateCtrlEnd;
	CDateTimeCtrl m_dateCtrlEndTime;

	sqlite3 * m_DB;

	CListCtrl m_listctrShowMail;
	CButton m_RadioBnClock;
	CEdit m_ctlDownloadStatus;

	CProgressBar* m_pDlgBar;

	CButton m_btnDownloadMail;

	HANDLE m_hDownloadThread;

	int m_nRadioBn;
	int m_nStuCountPos;
	int m_nSureStuCountPos;
	int m_nSureStuNamePos;
	int m_nNotSureStuNamePos;
	int m_nSureStuCount;
	int m_nNotSureStuCount;

	BOOL m_bInquiryFlag;
	BOOL m_bOnlySaveOnce;
	BOOL m_bResaveFlag;

	// �������ݿ�
	BOOL ConnectDB();

	/*
	*	��ʼ��list�ؼ�
	*	-----------------------------------------------
	*	|ְ����id | ְ�������� | ������ | �����洢·��|
	*	-----------------------------------------------
	*/
	BOOL InitListCtr();

	// �����ݿ����Ա��
	// �Զ�����Ա��ID
	BOOL AddEmployee();

	// ���ʼ����ص�ָ���ļ���
	// ����ʼ������������Ƿ����������������������������
	BOOL DownloadMail();

	// ��ʾ�ʼ��������
	void ShowMailDownloadStatus(_bstr_t strFileName, _bstr_t strOwnerName, _bstr_t strFilePath);

	// �ж��ʼ������ǰ������ĸ�ǲ���CR
	BOOL IsCR(std::vector<std::string>  & strMailSubject, COleDateTime colMailSendTime);

	// ȷ���ʼ�������ʱ�䷶Χ
	BOOL DownloadTime(COleDateTime& colDateStart, COleDateTime& colDateEnd);

	// ȷ�����ɱ����ʱ�䷶Χ
	BOOL ReportTime(COleDateTime& colDateStart, COleDateTime& colDateEnd);

	BOOL CheckDateCreateDir(std::vector<std::string>  & strMailSubject, CString & strDir, COleDateTime ColSendTime);

	// ����ʼ���Ϣ���ж��ʼ��Ƿ��Ѿ���ȡ
	BOOL MailIsRecv(_bstr_t bstrFileName);

	// ��¼�ʼ���������Ϣ
	BOOL RecordMailInfo(_bstr_t RecvTime, _bstr_t strFileName, _bstr_t strOwnerName, _bstr_t strFilePath);

	// ��ʱ������ռ���Ϣ
	BOOL ListMailInfo(COleDateTime & STime, COleDateTime & ETime);

	/*
	*	�����ָ����������ַ�����ȡ��Ϣ,ͬʱ����ʼ������Ϸ���
	*	�Ϸ������棬�Ƿ����ؼ�
	*/
	BOOL SplitFileName(std::vector<std::string> & Tag_pices, std::string & strFileName);

	/*
	*	��ѯְ�����������ְ��ID
	*	����Ҳ������������󷵻�<0
	*/
	int GetEmloyeeID(_bstr_t bstrEmpName);
	/*
	*	����Ŀ¼���ɱ���
	*/
	int GenStatment();

	// ��ѯ�ʼ�
	void Inquiry();

	// ��ȡ��ֹʱ��
	BOOL ReadDeadlineIni();

	afx_msg void OnBnClickedBnInquiry();
	afx_msg void OnBnClickedBnDownload();
	afx_msg void OnBnClickedBnAddstu();
	afx_msg void OnNMRClickListReportinfo(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnOperateDelstu();
protected:
	afx_msg LRESULT OnMainMsg(WPARAM wParam, LPARAM lParam);
public:
	CButton m_btnInquiry;
	afx_msg void OnDtnDatetimechangeDatetimepickerStartdate2(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnDtnDatetimechangeDatetimepickerEnddate2(NMHDR *pNMHDR, LRESULT *pResult);
};
