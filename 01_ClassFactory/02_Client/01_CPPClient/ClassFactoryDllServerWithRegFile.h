#pragma once
#include"unknwn.h"

class ISum :public IUnknown
{
public :
	//Isum specific method declaration
	virtual HRESULT __stdcall SumOfTwoIntegers(int, int, int*) = 0;//pure virtual
};

class ISubtract :public IUnknown
{
public:
	//ISubtract specific method declaration
	virtual HRESULT __stdcall SubtractionOfTwoIntegers(int, int, int*) = 0;//pure virtual
};

//CLSID of SumSubtarct Component {C7BD0DFB-03D1-4F1E-ABF2-D6C9C721A9D8}
const CLSID CLSID_SumSubtract =
{ 0xc7bd0dfb, 0x3d1, 0x4f1e, 0xab, 0xf2, 0xd6, 0xc9, 0xc7, 0x21, 0xa9, 0xd8 };

//IID of ISum Interface
const IID IID_ISum =
{ 0x2685eb90, 0xb191, 0x4cbc, 0xbc, 0xa1, 0x13, 0x9c, 0xa9, 0x5, 0xd1, 0x7c };

//IID of ISubtract Interface
const IID IID_ISubtract =
{ 0x93f622af, 0x49e8, 0x4b3b, 0x9f, 0x99, 0xa3, 0x81, 0xfd, 0x58, 0x2c, 0x94 };



