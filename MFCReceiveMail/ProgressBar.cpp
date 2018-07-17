// ProgressBar.cpp : 实现文件
//

#include "stdafx.h"
#include "MFCReceiveMail.h"
#include "ProgressBar.h"
#include "afxdialogex.h"

// CProgressBar 对话框

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


// CProgressBar 消息处理程序


BOOL CProgressBar::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  在此添加额外的初始化
	PostMessage(WM_LBUTTONDOWN);

	return TRUE;  // return TRUE unless you set the focus to a control
	// 异常:  OCX 属性页应返回 FALSE
}


void CProgressBar::OnLButtonDown(UINT nFlags, CPoint point)
{
	// TODO:  在此添加消息处理程序代码和/或调用默认值
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
	// TODO:  在此添加消息处理程序代码和/或调用默认值

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
