
// write_excelDlg.cpp : ʵ���ļ�
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

// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// Cwrite_excelDlg �Ի���



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


// Cwrite_excelDlg ��Ϣ�������

BOOL Cwrite_excelDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void Cwrite_excelDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR Cwrite_excelDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

const int TOP_FOLDER = 0;
const int CUSTOMER_FOLDER = 1;
const int DATE_FOLDER = 2;

static int folder_deep = -1;//��ʼ��-1����һ�ν���ѡ��Ķ���Ŀ¼TOP_FOLDER���0

void Cwrite_excelDlg::recurse_find_file( CString filePath,bool only_one_cus)
{
	folder_deep++;

	CFileFind fileFinder;

	//רΪ���һ�ҹ�˾ʹ��
	if(folder_deep == 1 && only_one_cus)
	{
		int back_slant_Pos=filePath.ReverseFind('\\');
		CString last_folder=filePath.Right(filePath.GetLength()-back_slant_Pos-1);
		_excel_data.set_current_customer(last_folder);
		_excel_data.add_customer(last_folder);
	}
	//רΪ���һ�ҹ�˾ʹ��

	filePath += "\\*.*";
	BOOL bFinished = fileFinder.FindFile(filePath);

	vector<detailed> vec;
	while(bFinished)
	{
		bFinished = fileFinder.FindNextFile();
		if(fileFinder.IsDirectory() && !fileFinder.IsDots())
	    {
			CString folder_name = fileFinder.GetFileName();
			if(folder_deep == 0)//�ҵ����ǿͻ��ļ���
			{
				_excel_data.set_current_customer(folder_name);
				_excel_data.add_customer(folder_name);
			}
			else if(folder_deep == 1)//�ҵ����������ļ���
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
				
				//28�Ű佱��ᱳ�� 15.2mx6.75mx3�� ���
				//find first blank
				int first_blank_pos = fileName.Find(_T(" "));
				//find firtst x or X
				int first_x_pos(-1),seconde_x_pos(-1);
				first_x_pos = fileName.Find(_T("x"),first_blank_pos+1);
				int upper_first_x = fileName.Find(_T("X"),first_blank_pos+1);
				//һ����дһ��Сд��ȡǰ���
				if((-1 != first_x_pos)&&(-1 != upper_first_x) && (first_x_pos > upper_first_x))
				{
					first_x_pos = upper_first_x;
				}
				//�������Ǵ�дX
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
					AfxMessageBox(error_file + "\n\n�ļ����Ƹ�ʽ���ԣ��������ļ��������ʽ������"); 
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
    bi.lpszTitle = _T("��ѡ����Ҫ�����Ŀ¼��");   
    bi.ulFlags = 0;   
    bi.lpfn = NULL;   
    bi.lParam = 0;   
    bi.iImage = 0;   
    //����ѡ��Ŀ¼�Ի���
    LPITEMIDLIST lp = SHBrowseForFolder(&bi);   

    if(lp && SHGetPathFromIDList(lp, szPath))   
    {
        //filePath.Format(_T("ѡ���Ŀ¼Ϊ %s"),  szPath);
        //AfxMessageBox(filePath); 
    }
    else   
	{
		AfxMessageBox(_T("��Ч��Ŀ¼��������ѡ��"));   
		return;
	}
        
	//�ݹ����Ŀ¼
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
    bi.lpszTitle = _T("��ѡ����Ҫ�����Ŀ¼��");   
    bi.ulFlags = 0;   
    bi.lpfn = NULL;   
    bi.lParam = 0;   
    bi.iImage = 0;   
    //����ѡ��Ŀ¼�Ի���
    LPITEMIDLIST lp = SHBrowseForFolder(&bi);   

    if(lp && SHGetPathFromIDList(lp, szPath))   
    {
        filePath.Format(_T("ѡ���Ŀ¼Ϊ %s"),  szPath);
        AfxMessageBox(filePath); 
    }
    else   
	{
		AfxMessageBox(_T("��Ч��Ŀ¼��������ѡ��"));   
		return;
	}
        
	//�ݹ����Ŀ¼
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
