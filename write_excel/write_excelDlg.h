
// write_excelDlg.h : ͷ�ļ�
//

#pragma once
#include "excel_data.h"
#include "excel_tool.h"

// Cwrite_excelDlg �Ի���
class Cwrite_excelDlg : public CDialogEx
{
// ����
public:
	Cwrite_excelDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_WRITE_EXCEL_DIALOG };

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
	afx_msg void OnBnClickedButton1();

private:
	void recurse_find_file( CString filePath, bool only_one_cus = false);
	void clear();
	bool is_valid_ext(const CString& ext);
private:
	excel_data _excel_data;	
public:
	afx_msg void OnBnClickedButton2();
};
