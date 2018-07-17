#pragma once
#include "afxcmn.h"
#include "afxwin.h"


// CProgressBar �Ի���

class CProgressBar : public CDialogEx
{
	DECLARE_DYNAMIC(CProgressBar)

public:
	CProgressBar(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~CProgressBar();

// �Ի�������
	enum { IDD = IDD_ProgressBar };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
//	afx_msg void OnTimer(UINT_PTR nIDEvent);

	CProgressCtrl m_ProgressCtl;

	virtual BOOL OnInitDialog();
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnClose();
};
