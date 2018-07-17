// ProgressBar.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "MFCReceiveMail.h"
#include "ProgressBar.h"
#include "afxdialogex.h"

// CProgressBar �Ի���

IMPLEMENT_DYNAMIC(CProgressBar, CDialogEx)

CProgressBar::CProgressBar(CWnd* pParent /*=NULL*/)
	: CDialogEx(CProgressBar::IDD, pParent)
{
	
}

CProgressBar::~CProgressBar()
{
}

void CProgressBar::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CProgressBar, CDialogEx)
	ON_WM_LBUTTONDOWN()
	ON_WM_CLOSE()
END_MESSAGE_MAP()


// CProgressBar ��Ϣ�������


BOOL CProgressBar::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  �ڴ���Ӷ���ĳ�ʼ��
	PostMessage(WM_LBUTTONDOWN);

	return TRUE;  // return TRUE unless you set the focus to a control
	// �쳣:  OCX ����ҳӦ���� FALSE
}


void CProgressBar::OnLButtonDown(UINT nFlags, CPoint point)
{
	// TODO:  �ڴ������Ϣ�����������/�����Ĭ��ֵ
	CRect rc;
	GetClientRect(&rc);

	CRect rcBar;
	rcBar.TopLeft().x = rc.Width() / 2 - 150;
	rcBar.TopLeft().y = rc.Height() / 2 - 20;
	rcBar.BottomRight().x = rc.Width() / 2 + 150;
	rcBar.BottomRight().y = rc.Height() / 2;

	m_ProgressCtl.Create(WS_CHILD | WS_VISIBLE | PBS_MARQUEE, rcBar, this, 11);
	m_ProgressCtl.SetMarquee(TRUE, 30);

	CDialogEx::OnLButtonDown(nFlags, point);
}


void CProgressBar::OnClose()
{
	// TODO:  �ڴ������Ϣ�����������/�����Ĭ��ֵ

// 	HWND hWnd = ::FindWindow(NULL, _T("MFCReceiveMail"));
// 
// 	::PostMessage(hWnd, BN_CLICKED, NULL, NULL);

	HWND hWnd = this->GetParent()->GetSafeHwnd();
	if (hWnd == NULL)
	{
		return;
	}

	::SendNotifyMessage(hWnd, WM_Main_MSG, NULL, NULL);

	CDialogEx::OnClose();
}
