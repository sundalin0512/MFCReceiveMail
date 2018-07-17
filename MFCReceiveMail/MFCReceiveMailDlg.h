
// MFCReceiveMailDlg.h : 头文件
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

// CMFCReceiveMailDlg 对话框
class CMFCReceiveMailDlg : public CDialogEx
{
// 构造
public:
	CMFCReceiveMailDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_MFCRECEIVEMAIL_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
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
	// 下载、查询班级
	CString m_cstrDownloadStatus;
	CString m_cstrClass;// 要添加学员所在的班级
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

	// 连接数据库
	BOOL ConnectDB();

	/*
	*	初始化list控件
	*	-----------------------------------------------
	*	|职工号id | 职工号姓名 | 附件名 | 附件存储路径|
	*	-----------------------------------------------
	*/
	BOOL InitListCtr();

	// 向数据库添加员工
	// 自动分配员工ID
	BOOL AddEmployee();

	// 将邮件下载到指定文件夹
	// 检查邮件附件的名字是否符合条件，符合条件就下载下来
	BOOL DownloadMail();

	// 显示邮件下载情况
	void ShowMailDownloadStatus(_bstr_t strFileName, _bstr_t strOwnerName, _bstr_t strFilePath);

	// 判断邮件主题的前两个字母是不是CR
	BOOL IsCR(std::vector<std::string>  & strMailSubject, COleDateTime colMailSendTime);

	// 确定邮件的下载时间范围
	BOOL DownloadTime(COleDateTime& colDateStart, COleDateTime& colDateEnd);

	// 确定生成报表的时间范围
	BOOL ReportTime(COleDateTime& colDateStart, COleDateTime& colDateEnd);

	BOOL CheckDateCreateDir(std::vector<std::string>  & strMailSubject, CString & strDir, COleDateTime ColSendTime);

	// 检查邮件信息，判断邮件是否已经收取
	BOOL MailIsRecv(_bstr_t bstrFileName);

	// 记录邮件附件的信息
	BOOL RecordMailInfo(_bstr_t RecvTime, _bstr_t strFileName, _bstr_t strOwnerName, _bstr_t strFilePath);

	// 按时间遍历收件信息
	BOOL ListMailInfo(COleDateTime & STime, COleDateTime & ETime);

	/*
	*	用来分附件的名字字符串提取信息,同时检查邮件命名合法性
	*	合法返回真，非法返回假
	*/
	BOOL SplitFileName(std::vector<std::string> & Tag_pices, std::string & strFileName);

	/*
	*	查询职工姓名，获得职工ID
	*	如果找不到，或发生错误返回<0
	*/
	int GetEmloyeeID(_bstr_t bstrEmpName);
	/*
	*	遍历目录生成报表
	*/
	int GenStatment();

	// 查询邮件
	void Inquiry();

	// 读取截止时间
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
