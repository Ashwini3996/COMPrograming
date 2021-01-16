#include<windows.h>
#include<stdio.h>
#include "AutomationServerWithRegFile.h"

//class declarations
//coClass

class CMyMath : public IMyMath
{
private : 
	long m_cRef;
	ITypeInfo* m_pITypeInfo = NULL;

public :
	//Constructor
	CMyMath(void);

	//Destructor
	~CMyMath(void);

	//IUnknown specific methode declaration (inherited)
	HRESULT __stdcall QueryInterface(REFIID, void**);
	ULONG __stdcall AddRef(void);
	ULONG __stdcall Release(void);

	//IDispatch specific methode declaration (inherited)
	HRESULT __stdcall GetTypeInfoCount(UINT*);
	HRESULT __stdcall GetTypeInfo(UINT, LCID, ITypeInfo**);
	HRESULT __stdcall GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*);
	HRESULT __stdcall Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*);

	//ISum specific methode declaration (inherited)
	HRESULT __stdcall SumOfTwoIntegers(int, int, int*);

	//ISubtract specific methode declaration (inherited)
	HRESULT __stdcall SubtractionOfTwoIntegers(int, int, int*);

	//custom methods
	HRESULT InitInstance(void);
};

//class factory
class CMyMathClassFactory : public IClassFactory
{
private :
	long m_cRef;

public :
	//Constuctor
	CMyMathClassFactory();

	//Destructor
	~CMyMathClassFactory();

    //IUnknown specific methode declaration (inherited)
    HRESULT __stdcall QueryInterface(REFIID, void**);
	ULONG __stdcall AddRef(void);
	ULONG __stdcall Release(void);

	//IClassFactory  specific methode declaration (inherited)
	HRESULT __stdcall CreateInstance(IUnknown*, REFIID, void**);
	HRESULT __stdcall LockServer(BOOL);    
};

//global veriable declarations
long glNumberOfActiveComponants = 0;//number of active componats
long glNumberOfServerLocks = 0;     //number of locks on this dll
const GUID LIBID_AutomationServer = { 0x4939d253, 0xae3a, 0x4d39, 0xaf, 0x5d, 0x35, 0x91, 0x92, 0x3f, 0x9c, 0xec };

//DllMain
BOOL WINAPI DllMain(HINSTANCE hDll, DWORD dwReason, LPVOID Reserved)
{
   //code
	switch (dwReason)
	{
		case DLL_PROCESS_ATTACH :
		break;    
		
		case DLL_THREAD_ATTACH :
		break;
	
		case DLL_THREAD_DETACH:
		break;
    
		case DLL_PROCESS_DETACH :
		break;
	}
	return TRUE;
}

CMyMath::CMyMath(void)
{
	m_cRef = 1;
	InterlockedIncrement(&glNumberOfActiveComponants);
}

CMyMath ::~CMyMath(void)
{
	InterlockedDecrement(&glNumberOfActiveComponants);
}

//Implementation CMyMath's  IUnknown's Method
HRESULT CMyMath::QueryInterface(REFIID riid, void** ppv)
{
	if (riid == IID_IUnknown)
		*ppv = static_cast<IMyMath*>(this);
	else if (riid == IID_IDispatch)
		*ppv = static_cast<IMyMath*>(this);
	else if (riid == IID_IMyMath)
		*ppv = static_cast<IMyMath*>(this);
	else
	{
		*ppv = NULL;
		return (E_NOINTERFACE);
	}

	reinterpret_cast<IUnknown*>(*ppv)->AddRef();
	return(S_OK);
}

ULONG CMyMath::AddRef(void)
{
	InterlockedIncrement(&m_cRef);
	return(m_cRef);
}

ULONG CMyMath::Release(void)
{
	InterlockedDecrement(&m_cRef);
	if (m_cRef == 0)
	{
		m_pITypeInfo->Release();
		m_pITypeInfo = NULL;
		delete(this);
		return 0;
	}
	return(m_cRef);
}

//Implementation of IMyMath's Method
HRESULT CMyMath::SumOfTwoIntegers(int num1, int num2, int* pSum)
{
	*pSum = num1 + num2;
	return(S_OK);
}

HRESULT CMyMath::SubtractionOfTwoIntegers(int num1, int num2, int* pSubtract)
{
	*pSubtract = num1 - num2;
	return(S_OK);
}

HRESULT CMyMath::InitInstance(void)
{
	void ComErrorDescriptionString(HWND, HRESULT);
	HRESULT hr;
	ITypeLib* pITypeLib = NULL;

	if (m_pITypeInfo == NULL)
	{
		hr = LoadRegTypeLib(LIBID_AutomationServer,
			1, 0,//major/minor version
			0x00,//LANG_NEUTRAL
			&pITypeLib);

		if (FAILED(hr))
		{
			ComErrorDescriptionString(NULL, hr);
			return hr;
		}

		hr = pITypeLib->GetTypeInfoOfGuid(IID_IMyMath, &m_pITypeInfo);
		if (FAILED(hr))
		{
			ComErrorDescriptionString(NULL, hr);
			pITypeLib->Release();
			return hr;
		}
		pITypeLib->Release();
	}
	return(S_OK);
}

 //Implentation of CMyMathClassFactory's Counstructor Method
CMyMathClassFactory::CMyMathClassFactory(void)
{
	m_cRef = 1;
}

//Implentation of CMyMathClassFactory's Destructor Method
CMyMathClassFactory::~CMyMathClassFactory(void)
{
}

//Implentation of CMyMathClassFactory's IClassFactory's IUnknown's Method
HRESULT CMyMathClassFactory::QueryInterface(REFIID riid, void** ppv)
{
	if (riid == IID_IUnknown)
		*ppv = static_cast<IClassFactory *>(this);
	else if (riid == IID_IClassFactory)
		*ppv = static_cast<IClassFactory *>(this);
	else
	{
		*ppv = NULL;
		return(E_NOINTERFACE);
	}
	reinterpret_cast<IUnknown*>(*ppv)->AddRef();
	return(S_OK);
}

ULONG CMyMathClassFactory::AddRef(void)
{
	InterlockedIncrement(&m_cRef);
	return(m_cRef);
}

ULONG CMyMathClassFactory::Release(void)
{
	InterlockedDecrement(&m_cRef);
	if (m_cRef == 0)
	{
		delete(this);
		return(0);
	}
	return(m_cRef);
}

HRESULT CMyMathClassFactory::CreateInstance(IUnknown* pUnkOuter, REFIID  riid, void** ppv)
{
	CMyMath* pCMyMath = NULL;
	HRESULT hr;

	if (pUnkOuter != NULL)
		return(CLASS_E_NOAGGREGATION);

	//create the instance of componant i.e. of CMyMath class
	pCMyMath = new CMyMath;
	
	if(pCMyMath == NULL)
		return(E_OUTOFMEMORY);
    
	//call automation related init method
	pCMyMath->InitInstance();

	//get the requested Interface
	hr = pCMyMath->QueryInterface(riid, ppv);

	pCMyMath->Release();
	return(hr);
}

HRESULT CMyMathClassFactory::LockServer(BOOL flock)
{
	if (flock)
		InterlockedIncrement(&glNumberOfServerLocks);
	else
		InterlockedDecrement(&glNumberOfServerLocks);
	return(S_OK);
}

//Implementation of IDispatch's method
HRESULT CMyMath::GetTypeInfoCount(UINT* pCountTypeInfo)
{	
   //as we have type library it is 1 else 0;
	*pCountTypeInfo = 1;
	return(S_OK);
}
HRESULT CMyMath::GetTypeInfo(UINT iTypeInfo, LCID lcid, ITypeInfo** ppITypeInfo)
{
	*ppITypeInfo = NULL;
	if (iTypeInfo != 0)
		return (DISP_E_BADINDEX);

	m_pITypeInfo->AddRef();
	*ppITypeInfo = m_pITypeInfo;
	return(S_OK);
}

HRESULT CMyMath::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
	return(DispGetIDsOfNames(m_pITypeInfo, rgszNames, cNames, rgDispId));
}

HRESULT CMyMath::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlag, DISPPARAMS* pDispParams,
	VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	HRESULT hr;
	hr = DispInvoke(this,
			m_pITypeInfo,
		dispIdMember,
		wFlag,
		pDispParams,
		pVarResult,
		pExcepInfo,
		puArgErr);
	return(hr);
}
	
//Implimentation of Exported Function From this Dll
HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
	CMyMathClassFactory* pCMyMathClassFactory = NULL;
	HRESULT hr;

	if (rclsid != CLSID_MyMath)
		return(CLASS_E_CLASSNOTAVAILABLE);

	//create class factory
	pCMyMathClassFactory = new CMyMathClassFactory;

	if (pCMyMathClassFactory == NULL)
		return(E_OUTOFMEMORY);

	hr = pCMyMathClassFactory->QueryInterface(riid, ppv);
	pCMyMathClassFactory->Release();
	return hr;
}


HRESULT __stdcall DllCanUnloadNow(void)
{
	if ((glNumberOfActiveComponants == 0) && (glNumberOfServerLocks == 0))
		return(S_OK);
	else
		return(S_FALSE);
}

void ComErrorDescriptionString(HWND hwnd, HRESULT hr)
{
	TCHAR* szErrorMessage = NULL;
	TCHAR str[255];

	if (FACILITY_WINDOWS == HRESULT_FACILITY(hr))
		hr = HRESULT_CODE(hr);

	if (
		FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
						NULL,
						hr,
						MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
						(LPTSTR)&szErrorMessage,
						0,
						NULL) != 0)
	{
		swprintf_s(str, TEXT("%#x : %s"), hr, szErrorMessage);
		LocalFree(szErrorMessage);
	}
	else
	{
		swprintf_s(str, TEXT("[Could not find a description for error # % #x.]\n"), hr);
		MessageBox(hwnd, str, TEXT("COM Error"), MB_OK);
	}
}
