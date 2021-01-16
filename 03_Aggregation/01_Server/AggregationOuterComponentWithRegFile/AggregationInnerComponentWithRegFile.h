#pragma once
#include"unknwn.h"

class IMultiplication :public IUnknown
{
public:
	//IMultiplication specific method declaration
	virtual HRESULT __stdcall MultiplicationOfTwoIntegers(int, int, int*) = 0;//pure virtual
};

class IDivision : public IUnknown
{
public:
	//IDivision specific method declaration
	virtual HRESULT __stdcall DivisionOfTwoIntegers(int, int, int*) = 0;//pure virtual
};

//CLSID of MultiplicationDivision Component {71ED6F50-E224-4A1F-AD73-59A6AF8517E7}
const CLSID CLSID_MultiplicationDivision =
{ 0x71ed6f50, 0xe224, 0x4a1f, 0xad, 0x73, 0x59, 0xa6, 0xaf, 0x85, 0x17, 0xe7 };

//IID of IMultiplication Interface
const IID IID_IMultiplication =
{ 0x7cf26fd4, 0xbe79, 0x46b7, 0x9e, 0x5c, 0xb1, 0xb5, 0xf3, 0x42, 0x6d, 0x3b };

//IID of IDivision Interface
const IID IID_IDivision =
{ 0x6ff6701e, 0x9e5d, 0x4f29, 0x97, 0xf, 0xd7, 0xc, 0xd7, 0xbb, 0x1f, 0xac };


