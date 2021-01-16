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

//CLSID of MultiplicationDivision Component {68A85301-4479-4F24-9C00-F7F06074D76D}
const CLSID CLSID_MultiplicationDivision =
{ 0x68a85301, 0x4479, 0x4f24, 0x9c, 0x0, 0xf7, 0xf0, 0x60, 0x74, 0xd7, 0x6d };

//IID of IMultiplication Interface
const IID IID_IMultiplication =
{ 0x961bef2, 0x2042, 0x4e52, 0x9e, 0xf1, 0xc0, 0xa, 0x57, 0x8, 0xa4, 0x1f };

//IID of IDivision Interface
const IID IID_IDivision =
{ 0xd1399e35, 0x54f4, 0x473c,  0xa8, 0x61, 0xcf, 0x9c, 0x94, 0x96, 0x12, 0x37 };

