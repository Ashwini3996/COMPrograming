#pragma once
#include"unknwn.h"
#undef INTERFACE
#define INTERFACE IMyMath

DECLARE_INTERFACE_(IMyMath, IDispatch)
{	
	STDMETHOD(SumOfTwoIntegers)(THIS_ int, int, int*) PURE;	
	STDMETHOD(SubtractionOfTwoIntegers)(THIS_ int, int, int*) PURE;		
};

//CLSID of MyMath Component : {CDCB1DA6-7BB3-4A2B-91CE-E963E27FB3DC}
const CLSID CLSID_MyMath = { 0xcdcb1da6, 0x7bb3, 0x4a2b, 0x91, 0xce, 0xe9, 0x63, 0xe2, 0x7f, 0xb3, 0xdc };

//IID of ISum Interface : {8C8B9461-21EF-4A7D-94F2-C0B1BAD0A1F7}
const IID IID_IMyMath = { 0x8c8b9461, 0x21ef, 0x4a7d, 0x94, 0xf2, 0xc0, 0xb1, 0xba, 0xd0, 0xa1, 0xf7 };
