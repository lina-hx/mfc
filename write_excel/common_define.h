#ifndef _COMMON_DEFINE_H_
#define _COMMON_DEFINE_H_
#include <CString>
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

	CString name;
	double length;
	double height;
	unsigned int count;
	double area;
	double unit_price;
	double total_price;
};
#endif