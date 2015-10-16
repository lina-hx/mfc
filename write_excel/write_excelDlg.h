
// write_excelDlg.h : 头文件
//

#pragma once
#include "excel_data.h"
#include "excel_tool.h"

// Cwrite_excelDlg 对话框
class Cwrite_excelDlg : public CDialogEx
{
// 构造
public:
	Cwrite_excelDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_WRITE_EXCEL_DIALOG };

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
	afx_msg void OnBnClickedButton1();

private:
	void recurse_find_file( CString filePath);
private:
	excel_data _excel_data;	
	excel_tool _excel_tool;
public:
	afx_msg void OnBnClickedButton2();
};
