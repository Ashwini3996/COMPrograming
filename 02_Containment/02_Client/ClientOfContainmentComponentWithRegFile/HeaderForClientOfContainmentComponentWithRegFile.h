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

//CLSID of SumSubtarct Component {E1E17BEF-498F-4BE0-AF0A-E6AFB44634B1}
const CLSID CLSID_SumSubtract =
{ 0xe1e17bef, 0x498f, 0x4be0, 0xaf, 0xa, 0xe6, 0xaf, 0xb4, 0x46, 0x34, 0xb1 };

//IID of ISum Interface
const IID IID_ISum =
{ 0x28beb688, 0x328a, 0x40f9, 0xb4, 0x36, 0x75, 0x14, 0x91, 0xe, 0x12, 0x20 };

//IID of ISubtract Interface
const IID IID_ISubtract =
{ 0x3a6aa69e, 0xb99f, 0x4c5d, 0x98, 0xa9, 0xa2, 0x93, 0x42, 0x12, 0x72, 0x5a };

//IID of IMultiplication Interface
const IID IID_IMultiplication =
{ 0x961bef2, 0x2042, 0x4e52, 0x9e, 0xf1, 0xc0, 0xa, 0x57, 0x8, 0xa4, 0x1f };

//IID of IDivision Interface
const IID IID_IDivision =
{ 0xd1399e35, 0x54f4, 0x473c,  0xa8, 0x61, 0xcf, 0x9c, 0x94, 0x96, 0x12, 0x37 };
