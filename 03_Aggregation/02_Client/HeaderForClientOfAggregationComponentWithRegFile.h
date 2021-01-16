#pragma once
#include"unknwn.h"

class ISum :public IUnknown
{
public:
	//Isum specific method declaration
	virtual HRESULT __stdcall SumOfTwoIntegers(int, int, int*) = 0;//pure virtual
};

class ISubtract :public IUnknown
{
public:
	//ISubtract specific method declaration
	virtual HRESULT __stdcall SubtractionOfTwoIntegers(int, int, int*) = 0;//pure virtual
};

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


//CLSID of SumSubtarct Component  {015C91D2-1A38-45B1-A32B-9180CFD11CC8}
const CLSID CLSID_SumSubtract =
{ 0x15c91d2, 0x1a38, 0x45b1, 0xa3, 0x2b, 0x91, 0x80, 0xcf, 0xd1, 0x1c, 0xc8 };

//IID of ISum Interface
const IID IID_ISum =
{ 0x3dc2fbb6, 0xc005, 0x4320, 0x9b, 0xae, 0xef, 0x4b, 0x55, 0x1c, 0x42, 0xdd };

//IID of ISubtract Interface
const IID IID_ISubtract =
{ 0x14d1dca0, 0x88c1, 0x4df3, 0xb7, 0x86, 0x39, 0xf2, 0xbc, 0xe9, 0x3a, 0x27 };

//IID of IMultiplication Interface
const IID IID_IMultiplication =
{ 0x7cf26fd4, 0xbe79, 0x46b7, 0x9e, 0x5c, 0xb1, 0xb5, 0xf3, 0x42, 0x6d, 0x3b };

//IID of IDivision Interface
const IID IID_IDivision =
{ 0x6ff6701e, 0x9e5d, 0x4f29, 0x97, 0xf, 0xd7, 0xc, 0xd7, 0xbb, 0x1f, 0xac };

