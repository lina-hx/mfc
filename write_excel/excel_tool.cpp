#include "stdafx.h"
#include "excel_tool.h"

CApplication excel_tool::_app;
	CFont0 excel_tool::_font;
	CRange excel_tool::_range;
	CRange excel_tool::_range_merge;
	CRange excel_tool::_cols;
	CWorkbook excel_tool::_book;
	CWorkbooks excel_tool::_books;
	CWorksheet excel_tool::_sheet;
	CWorksheets excel_tool::_sheets;

	unsigned int excel_tool::_current_row = 1; 