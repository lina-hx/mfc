
// write_excel.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// Cwrite_excelApp:
// �йش����ʵ�֣������ write_excel.cpp
//

class Cwrite_excelApp : public CWinApp
{
public:
	Cwrite_excelApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern Cwrite_excelApp theApp;