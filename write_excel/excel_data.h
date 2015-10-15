#ifndef _EXCEL_DATA_H_
#define _EXCEL_DATA_H_

#include <map>
#include <string>
#include <vector>
using namespace std;

struct detailed
{
	detailed()
	{
		name = "";
		length = 0;
		height = 0;
		count = 0;
		area = 0;
		unit_price = 0;
		total_price = 0;
	}

	string name;
	unsigned int length;
	unsigned int height;
	unsigned int count;
	unsigned int area;
	unsigned int unit_price;
	unsigned int total_price;
};

typedef map<CString,map<CString,vector<detailed> > >::iterator _all_data_it;
typedef map<CString,vector<detailed> >::iterator _one_data_it;

class excel_data
{
public:

	void add_customer_and_one_day(const CString& customer, const CString& date)
	{
		// a empty vector
		vector<detailed> vec;

		map<CString,vector<detailed> > tmp;
		tmp.insert(make_pair(date,vec));

		_all_data_map.insert(make_pair(customer,tmp));
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