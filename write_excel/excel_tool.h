#ifndef _EXCEL_TOOL_H_
#define _EXCEL_TOOL_H_

#include "CApplication.h"
#include "CFont0.h"
#include "CRange.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include "common_define.h"

static CString g_excel_column[27] = {"","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T"\
"U","V","W","X","Y","Z"};

static CString  g_header[9] = {"","时间","文件名","长","高","数量","面积","单价","金额"};

class excel_tool
{
public:

	static bool init()
	{
		if (!AfxOleInit())
		{   
			AfxMessageBox(_T("AfxOleInit error"));
			return FALSE;
		}

		COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
		if (!_app.CreateDispatch(_T("Excel.Application")))
		{   
			AfxMessageBox(_T("无法创建Excel应用！")); 
			return TRUE;  
		}
		_books = _app.get_Workbooks();
		_book = _books.Add(covOptional);
		_sheets = _book.get_Worksheets();
		_sheet = _sheets.get_Item(COleVariant((short)1));
		_range = _sheet.get_Range(COleVariant(_T("A1")),COleVariant(_T("D10")));
		
		//_range.put_Value2(COleVariant(_T("Hello Excel")));
		
		_range.put_Item(COleVariant((long)1),COleVariant((long)1),COleVariant(_T("bb Excel asdasd asdas 1231 asd 1 123 ")));
		_cols = _range.get_EntireColumn();
		_cols.AutoFit();

		_range_merge = _sheet.get_Range(COleVariant(_T("A1")),COleVariant(_T("A10")));
		_range_merge.Merge(COleVariant((short)0));

		_app.put_Visible(true);
		_app.put_UserControl(true);
		//_book.SaveCopyAs(COleVariant(_T("c:/ha123.xlsx")));
		//_book.put_Saved(true);

		_book.ReleaseDispatch();
		_books.ReleaseDispatch();
		_app.ReleaseDispatch();
		_app.Quit();

		return true;
	}

	static void init2()
	{
		if(!_b_init)
		{
			if (!AfxOleInit())
			{   
				AfxMessageBox(_T("AfxOleInit error"));
				return;
			}
			_b_init = true;
		}
		

		COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR); 
		if (!_app.CreateDispatch(_T("Excel.Application")))
		{   
			AfxMessageBox(_T("无法创建Excel应用！")); 
			return;  
		}
		_books = _app.get_Workbooks();
		_book = _books.Add(covOptional);
		_sheets = _book.get_Worksheets();
		_sheet = _sheets.get_Item(COleVariant((short)1));
	}

public:
	static void write_header(const CString& customer)
	{
		//客户名称
		_range_merge = _sheet.get_Range(COleVariant(_T("A1")),COleVariant(_T("H1")));
		_range_merge.Merge(COleVariant((short)0));
		_cols = _range_merge.get_EntireColumn();
		_cols.AutoFit();
		_font = _range_merge.get_Font();
		_font.put_Bold(COleVariant((short)TRUE));
		_range_merge.put_HorizontalAlignment(COleVariant((long)-4108));
		_range_merge.put_Item(COleVariant((long)1),COleVariant((long)1),COleVariant(customer));
		_current_row++;

		//时间 文件名 长 高 数量 面积 单价 金额
		_range = _sheet.get_Range(COleVariant(_T("A2")),COleVariant(_T("H2")));
		_cols = _range.get_EntireColumn();
		_cols.AutoFit();
	    _font = _range.get_Font();
		_font.put_Bold(COleVariant((short)TRUE));
		_range.put_HorizontalAlignment(COleVariant((long)-4108));

		for(int i=1; i <= 8; i++)
		{
			_range.put_Item(COleVariant((long)1),COleVariant((long)i),COleVariant(g_header[i]));
		}
		_current_row++;
		_app.put_Visible(true);
		_app.put_UserControl(true);
	}

	static void write_one_line_data(const CString& date,const detailed& d)
	{
		CString left_top,right_bottom;
		left_top.Format("A%d",_current_row);
		right_bottom.Format("H%d",_current_row);
		_range = _sheet.get_Range(COleVariant(left_top),COleVariant(right_bottom));
		_range.put_HorizontalAlignment(COleVariant((long)-4108));
		_cols = _range.get_EntireColumn();
		_cols.AutoFit();

		int i = 1;
		_range.put_Item(COleVariant((long)1),COleVariant((long)i++),COleVariant(date));
		_range.put_Item(COleVariant((long)1),COleVariant((long)i++),COleVariant(_T(d.name)));
		_range.put_Item(COleVariant((long)1),COleVariant((long)i++),COleVariant((double)d.length));
		_range.put_Item(COleVariant((long)1),COleVariant((long)i++),COleVariant((double)d.height));
		_range.put_Item(COleVariant((long)1),COleVariant((long)i++),COleVariant((long)d.count));
		_range.put_Item(COleVariant((long)1),COleVariant((long)i++),COleVariant((double)d.area));
		_range.put_Item(COleVariant((long)1),COleVariant((long)i++),COleVariant((double)d.unit_price));
		_range.put_Item(COleVariant((long)1),COleVariant((long)i++),COleVariant((double)d.total_price));
		
		// _range得到的范围只有一行，所以row索引就一行
		_cols = _range.get_Range(COleVariant("F1"),COleVariant(_T("F1")));
		CString len_col,height_col,count_col;
		len_col.Format("C%d",_current_row);
		height_col.Format("D%d",_current_row);
		count_col.Format("E%d",_current_row);
		CString formula("=");
		formula += len_col + "*" + height_col + "*" + count_col;
		_cols.put_Formula(COleVariant(formula));

		_cols = _range.get_Range(COleVariant("H1"),COleVariant(_T("H1")));
		CString area_col,unit_price_col;
		area_col.Format("F%d",_current_row);
		unit_price_col.Format("G%d",_current_row);
		formula.Format("=");
		formula += area_col + "*" + unit_price_col;
		_cols.put_Formula(COleVariant(formula));

		_current_row++;
		_app.put_Visible(true);
		_app.put_UserControl(true);
	}

	static void merge_same_date()
	{
		CString right_bottom;
		right_bottom.Format("A%d",_current_row);
		_range = _sheet.get_Range(COleVariant("A1"),COleVariant(right_bottom));
		
		unsigned int start_row(INT_MAX),end_row(0);

		for(unsigned int i = 1; i < _current_row; i++)
		{
			COleVariant d1 = _range.get_Item(COleVariant((long)i),COleVariant((long)1));
			COleVariant d2 = _range.get_Item(COleVariant((long)i+1),COleVariant((long)1));
			d1.ChangeType(VT_BSTR);
			d2.ChangeType(VT_BSTR);

			CString date1,date2;
			date1=d1.bstrVal;
			date2=d2.bstrVal;

			if(date1 == date2)
			{
				if(i < start_row)
				{
					start_row = i;
				}
				if(i+1 > end_row)
				{
					end_row = i+1;
				}
			}
			else if(date1 != date2)
			{
				if(start_row != INT_MAX && end_row != 0)
				{
					CString col1,col2;
					col1.Format("A%d",start_row);
					col2.Format("A%d",end_row);
					_range_merge =_sheet.get_Range(COleVariant(col1),COleVariant(col2));
					_range_merge.Merge(COleVariant((short)0));

					start_row = INT_MAX;
					end_row = 0;
				}
			}
		}
	}

	static void close_excel()
	{
		_book.ReleaseDispatch();
		_books.ReleaseDispatch();
		_app.ReleaseDispatch();
		_app.Quit();
		_current_row = 1;
	}
private:
	static CApplication _app;
	static CFont0 _font;
	static CRange _range;
	static CRange _range_merge;
	static CRange _cols;
	static CWorkbook _book;
	static CWorkbooks _books;
	static CWorksheet _sheet;
	static CWorksheets _sheets;
	
	static unsigned int _current_row;
	static bool _b_init;
};
#endif