
// write_excelDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "write_excel.h"
#include "write_excelDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

const int one_word_bytes = 1;
const int one_chinese_word_bytes = 2;
const int ext_len = 6;
const CString ext_arry[6] = {".jpg",".jpeg",".bmp",".png",".tiff",".gif"}; 

// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
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


// Cwrite_excelDlg 对话框



Cwrite_excelDlg::Cwrite_excelDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cwrite_excelDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void Cwrite_excelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(Cwrite_excelDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &Cwrite_excelDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &Cwrite_excelDlg::OnBnClickedButton2)
END_MESSAGE_MAP()


// Cwrite_excelDlg 消息处理程序

BOOL Cwrite_excelDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void Cwrite_excelDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void Cwrite_excelDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR Cwrite_excelDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

const int TOP_FOLDER = 0;
const int CUSTOMER_FOLDER = 1;
const int DATE_FOLDER = 2;

static int folder_deep = -1;//初始化-1，第一次进入选择的顶层目录TOP_FOLDER变成0

void Cwrite_excelDlg::recurse_find_file( CString filePath,bool only_one_cus)
{
	folder_deep++;

	CFileFind fileFinder;

	//专为输出一家公司使用
	if(folder_deep == 1 && only_one_cus)
	{
		int back_slant_Pos=filePath.ReverseFind('\\');
		CString last_folder=filePath.Right(filePath.GetLength()-back_slant_Pos-1);
		_excel_data.set_current_customer(last_folder);
		_excel_data.add_customer(last_folder);
	}
	//专为输出一家公司使用

	filePath += "\\*.*";
	BOOL bFinished = fileFinder.FindFile(filePath);

	vector<detailed> vec;
	while(bFinished)
	{
		bFinished = fileFinder.FindNextFile();
		if(fileFinder.IsDirectory() && !fileFinder.IsDots())
	    {
			CString folder_name = fileFinder.GetFileName();
			if(folder_deep == 0)//找到的是客户文件夹
			{
				_excel_data.set_current_customer(folder_name);
				_excel_data.add_customer(folder_name);
			}
			else if(folder_deep == 1)//找到的是日期文件夹
			{
				_excel_data.set_current_date(folder_name);
				_excel_data.add_one_day_for_customer(_excel_data.get_current_customer(),folder_name);
			}
			recurse_find_file(fileFinder.GetFilePath());
		}
		else
		{
			CString fileName = fileFinder.GetFileName();
            int dotPos=fileName.ReverseFind('.');
            CString fileExt=fileName.Right(fileName.GetLength()-dotPos);
			if(is_valid_ext(fileExt))	
			{
				bool format_error = false;
				
				//28号颁奖晚会背景 15.2mx6.75mx3张 喷绘
				//find first blank
				int first_blank_pos = fileName.Find(_T(" "));
				//find firtst x or X
				int first_x_pos(-1),seconde_x_pos(-1);
				first_x_pos = fileName.Find(_T("x"),first_blank_pos+1);
				int upper_first_x = fileName.Find(_T("X"),first_blank_pos+1);
				//一个大写一个小写，取前面的
				if((-1 != first_x_pos)&&(-1 != upper_first_x) && (first_x_pos > upper_first_x))
				{
					first_x_pos = upper_first_x;
				}
				//两个都是大写X
				if(-1 == first_x_pos && upper_first_x != -1)
				{
					first_x_pos = upper_first_x;
				}

				//find second x or X
				seconde_x_pos = fileName.Find(_T("x"),first_x_pos+1);
				if(-1 == seconde_x_pos)
				{
					seconde_x_pos = fileName.Find(_T("X"),first_x_pos+1);
				}
				if(-1 == seconde_x_pos || -1 == first_x_pos)
				{
					CString error_file = filePath.Left(filePath.GetLength()-3) + fileName;
					AfxMessageBox(error_file + "\n\n文件名称格式不对，跳过此文件，请检查格式！！！"); 
					continue;
				}
				//find seconde blank
				int seconde_blank_pos = fileName.Find(_T(" "),seconde_x_pos+1);
				//int count_pos = fileName.Find(_T(" "),seconde_blank_pos+1);

				CString length = fileName.Mid(first_blank_pos+1,first_x_pos-first_blank_pos-1-one_word_bytes);
				CString height = fileName.Mid(first_x_pos+1,seconde_x_pos-first_x_pos-1-one_word_bytes);
				CString count = fileName.Mid(seconde_x_pos+1,seconde_blank_pos-seconde_x_pos-1-one_chinese_word_bytes);

				CString file = fileName.Left(dotPos);
				detailed record;
				record.name = file;
				record.length = _ttof(length);
				record.height = _ttof(height);
				record.count = _ttoi(count);

				vec.push_back(record);
			}
		}
	}
	fileFinder.Close();
	if(!vec.empty())
	{
		_excel_data.add_one_day_details(vec);
	}
	folder_deep--;
}

void Cwrite_excelDlg::clear()
{
	folder_deep = -1;
	_excel_data.clear();
}

void Cwrite_excelDlg::OnBnClickedButton1()
{
	clear();

	char szPath[MAX_PATH];
	CString filePath;

    ZeroMemory(szPath, sizeof(szPath));   

    BROWSEINFO bi;   
    bi.hwndOwner = m_hWnd;   
    bi.pidlRoot = NULL;   
    bi.pszDisplayName = szPath;   
    bi.lpszTitle = _T("请选择需要打包的目录：");   
    bi.ulFlags = 0;   
    bi.lpfn = NULL;   
    bi.lParam = 0;   
    bi.iImage = 0;   
    //弹出选择目录对话框
    LPITEMIDLIST lp = SHBrowseForFolder(&bi);   

    if(lp && SHGetPathFromIDList(lp, szPath))   
    {
        //filePath.Format(_T("选择的目录为 %s"),  szPath);
        //AfxMessageBox(filePath); 
    }
    else   
	{
		AfxMessageBox(_T("无效的目录，请重新选择"));   
		return;
	}
        
	//递归遍历目录
	filePath.Empty();
	filePath.Format(_T("%s"),  szPath);
	recurse_find_file(filePath);
	
	CEdit* p_edit = (CEdit *)this->GetDlgItem(IDC_EDIT1);
	CString tail;
	p_edit->GetWindowTextA(tail);
	_excel_data.set_tail(tail);

	_excel_data.output_all_customer_excel();
}


void Cwrite_excelDlg::OnBnClickedButton2()
{
	clear();

	char szPath[MAX_PATH];
	CString filePath;

    ZeroMemory(szPath, sizeof(szPath));   

    BROWSEINFO bi;   
    bi.hwndOwner = m_hWnd;   
    bi.pidlRoot = NULL;   
    bi.pszDisplayName = szPath;   
    bi.lpszTitle = _T("请选择需要打包的目录：");   
    bi.ulFlags = 0;   
    bi.lpfn = NULL;   
    bi.lParam = 0;   
    bi.iImage = 0;   
    //弹出选择目录对话框
    LPITEMIDLIST lp = SHBrowseForFolder(&bi);   

    if(lp && SHGetPathFromIDList(lp, szPath))   
    {
        filePath.Format(_T("选择的目录为 %s"),  szPath);
        AfxMessageBox(filePath); 
    }
    else   
	{
		AfxMessageBox(_T("无效的目录，请重新选择"));   
		return;
	}
        
	//递归遍历目录
	folder_deep++;

	filePath.Empty();
	filePath.Format(_T("%s"),  szPath);
	recurse_find_file(filePath,true);

	_excel_data.output_all_customer_excel();
}

bool Cwrite_excelDlg::is_valid_ext(const CString& ext)
{
	for(int i = 0; i < ext_len; i++)
	{
		if(ext == ext_arry[i])
		{
			return true;
		}
	}
	return false;
}
