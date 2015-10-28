
// write_excelDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "write_excel.h"
#include "write_excelDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


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

void Cwrite_excelDlg::recurse_find_file( CString filePath)
{
	folder_deep++;

	CFileFind fileFinder;
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
			if(fileExt == _T(".jpg"))
			{
				//28�Ű佱��ᱳ�� 15.2x6.75 60 ���.jpg
				int first_blank_pos = fileName.Find(_T(" "));
				int x_pos = fileName.Find(_T("x"),first_blank_pos+1);
				int seconde_blank_pos = fileName.Find(_T(" "),x_pos+1);
				int count_pos = fileName.Find(_T(" "),seconde_blank_pos+1);

				CString length = fileName.Mid(first_blank_pos+1,x_pos-first_blank_pos-1);
				CString height = fileName.Mid(x_pos+1,seconde_blank_pos-x_pos-1);
				CString count = fileName.Mid(seconde_blank_pos+1,count_pos-seconde_blank_pos-1);

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

void Cwrite_excelDlg::OnBnClickedButton1()
{
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
        AfxMessageBox(_T("��Ч��Ŀ¼��������ѡ��"));   

	//�ݹ����Ŀ¼
	filePath.Empty();
	filePath.Format(_T("%s"),  szPath);
	recurse_find_file(filePath);

	_excel_data.output_all_customer_excel();
}


void Cwrite_excelDlg::OnBnClickedButton2()
{
	//_excel_data
}
