#ifndef _EXCEL_DATA_H_
#define _EXCEL_DATA_H_

#include <map>
#include <string>
#include <vector>
#include "common_define.h"
#include "excel_tool.h"
using namespace std;

typedef map<CString,map<CString,vector<detailed> > >::iterator _all_data_it;
typedef map<CString,vector<detailed> >::iterator _one_data_it;

class excel_data
{
public:

	void add_customer(const CString& customer)
	{
		// a empty vector
		vector<detailed> vec;

		map<CString,vector<detailed> > tmp;

		_all_data_map.insert(make_pair(customer,tmp));
	}

	void add_one_day_for_customer(const CString& cus, const CString& date)
	{
		_all_data_it it = _all_data_map.find(cus);
		if(it == _all_data_map.end())
		{
			return;
		}
		map<CString,vector<detailed> >& tmp = it->second;

		vector<detailed> vec;
		tmp.insert(make_pair(date,vec));
	}

	void add_one_day_details(const vector<detailed>& vec)
	{
		_all_data_it it = _all_data_map.find(_current_customer);
		if(it == _all_data_map.end())
		{
			return;
		}
		map<CString,vector<detailed> >& tmp = it->second;
		_one_data_it it2 = tmp.find(_current_date);
		if(it2 == tmp.end())
		{
			return;
		}
		vector<detailed>& map_vec = it2->second;
		copy(vec.begin(),vec.end(),back_inserter(map_vec));
	}

	void output_all_customer_excel()
	{
		_all_data_it it = _all_data_map.begin();
		for(; it != _all_data_map.end(); it++)
		{
			output_one_customer_excel(it->first);
		}
	}

	void output_one_customer_excel(const CString& cus)
	{
		CString header = cus;
		header += "喷绘加工明细表";
		excel_tool::init2();
		excel_tool::write_header(header);

		_all_data_it it = _all_data_map.find(cus);
		if(it == _all_data_map.end())
		{
			return;
		}
		map<CString,vector<detailed> >& one_company = it->second;
		_one_data_it one_day_it = one_company.begin();
		for(;one_day_it != one_company.end(); one_day_it++)
		{
			vector<detailed>& detail_vec = one_day_it->second;
			for(vector<detailed>::const_iterator vec_it = detail_vec.begin(); vec_it!=detail_vec.end();vec_it++)
			{
				// write one line every time
				excel_tool::write_one_line_data(one_day_it->first,*vec_it);
			}
		}

		excel_tool::close_excel();
	}

	void set_current_customer(const CString& cus)
	{
		_current_customer = cus;
	}

	void set_current_date(const CString& date)
	{
		_current_date = date;
	}

	CString get_current_customer()
	{
		return _current_customer;
	}

	CString get_current_date()
	{
		return _current_date;
	}

private:
	//第一个CString是各个客户
	//第二个map<CString,vector<detailed>>表示这个客户下面所有日期的所有产品
	//CString表示不同日期
	//detailed表示一行明细
	map<CString,map<CString,vector<detailed> > > _all_data_map;

	CString _current_customer;
	CString _current_date;
};

#endif