
// MFCReceiveMail.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������

#define WM_Main_MSG WM_USER+0x01001


// CMFCReceiveMailApp: 
// �йش����ʵ�֣������ MFCReceiveMail.cpp
//

class CMFCReceiveMailApp : public CWinApp
{
public:
	CMFCReceiveMailApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CMFCReceiveMailApp theApp;