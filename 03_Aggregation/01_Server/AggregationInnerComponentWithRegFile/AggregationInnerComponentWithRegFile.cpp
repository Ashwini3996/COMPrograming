#include<windows.h>
#include"AggregationInnerComponentWithRegFile.h"

//interface declaration (for internal uesd only i.e not to be included in .h file)
interface INoAggregationIUnknown //new 
{
	virtual HRESULT __stdcall QueryInterface_NoAggregation(REFIID, void**) = 0;
	virtual ULONG __stdcall AddRef_NoAggregation(void) = 0;
	virtual ULONG __stdcall Release_NoAggregation(void) = 0;
};

//class declaration
class CMultiplicationDivision :public INoAggregationIUnknown, IMultiplication, IDivision
{
private:
	long m_cRef;
	IUnknown* m_pIUnknownOuter;

public:
	//constructor method declartion
	CMultiplicationDivision(IUnknown *);//new

	//destructor method declartion
	~CMultiplicationDivision(void);

	//IUnknown specific method declartion(inherited)
	HRESULT __stdcall QueryInterface(REFIID, void**);
	ULONG __stdcall   AddRef(void);
	ULONG __stdcall   Release(void);

	virtual HRESULT __stdcall QueryInterface_NoAggregation(REFIID, void**);//new
	virtual ULONG __stdcall AddRef_NoAggregation(void);//new
	virtual ULONG __stdcall Release_NoAggregation(void);//new

	//IMultiplication specific method declartion(inherited)
	HRESULT __stdcall MultiplicationOfTwoIntegers(int, int, int*);

	//IDivision specific method declartion(inherited)
	HRESULT __stdcall DivisionOfTwoIntegers(int, int, int*);
};

class CMultiplicationDivisionClassFactory :public IClassFactory
{
private:
	long m_cRef;

public:
	//constructor method declartion
	CMultiplicationDivisionClassFactory(void);

	//destructor method declartion
	~CMultiplicationDivisionClassFactory(void);

	//IUnknown specific method declartion(inherited)
	HRESULT __stdcall QueryInterface(REFIID, void**);
	ULONG __stdcall   AddRef(void);
	ULONG __stdcall   Release(void);

	//IClassFactory specific method declartion(inherited)
	HRESULT __stdcall CreateInstance(IUnknown*, REFIID, void**);
	HRESULT __stdcall LockServer(BOOL);
};

//global veriable declarations
long glNumberOfActiveComponants = 0;//number of active components
long glNumberOfServerLocks = 0; //number of locks on this Dll

//DllMain
BOOL WINAPI DllMain(HINSTANCE hDll, DWORD dwReason, LPVOID Reserved)
{
	//code
	switch (dwReason)
	{
	case DLL_PROCESS_ATTACH:
		break;

	case DLL_THREAD_ATTACH:
		break;

	case DLL_THREAD_DETACH:
		break;

	case DLL_PROCESS_DETACH:
		break;
	}
	return(TRUE);
}

//Implemetation of CMultiplicationDivision's Constructor method
CMultiplicationDivision::CMultiplicationDivision(IUnknown* pIUnknownOuter)
{
	m_cRef = 1; //hardcoded initialization to anticipate passible failuer of QueryInterface()
	InterlockedIncrement(&glNumberOfActiveComponants);//Increment global counter

	if (pIUnknownOuter != NULL)
		m_pIUnknownOuter = pIUnknownOuter;
	else
		m_pIUnknownOuter = reinterpret_cast <IUnknown*> 
		(static_cast<INoAggregationIUnknown *>(this));
}

//Implemetation of CMultiplicationDivision's Destructor method
CMultiplicationDivision::~CMultiplicationDivision(void)
{
	InterlockedDecrement(&glNumberOfActiveComponants);//Decrement global counter
}

//Implemetation of CMultiplicationDivision's method supporting IUnknown method
HRESULT CMultiplicationDivision::QueryInterface(REFIID riid, void** ppv)
{
	return(m_pIUnknownOuter->QueryInterface(riid, ppv));
}

ULONG CMultiplicationDivision::AddRef(void)
{
	return(m_pIUnknownOuter->AddRef());
}

ULONG CMultiplicationDivision::Release(void)
{
	return(m_pIUnknownOuter->Release());
}


//Implemetation of CMultiplicationDivision's Aggregation NonSupporting IUnknown's method
HRESULT CMultiplicationDivision::QueryInterface_NoAggregation(REFIID riid, void** ppv)
{
	//code
	if (riid == IID_IUnknown)
		*ppv = static_cast<INoAggregationIUnknown *>(this);
	else if (riid == IID_IMultiplication)
		*ppv = static_cast<IMultiplication*>(this);
	else if (riid == IID_IDivision)
		*ppv = static_cast<IDivision*>(this);
	else
	{
		*ppv = NULL;
		return(E_NOINTERFACE);
	}
	reinterpret_cast<IUnknown*>(*ppv)->AddRef();
	return(S_OK);
}

ULONG CMultiplicationDivision::AddRef_NoAggregation(void)
{
	//code
	InterlockedIncrement(&m_cRef);
	return(m_cRef);
}

ULONG CMultiplicationDivision::Release_NoAggregation(void)
{
	//code
	InterlockedDecrement(&m_cRef);
	if (m_cRef == 0)
	{
		delete(this);
		return(0);
	}
	return(m_cRef);
}

//Imlementation of IMultiplication's methods
HRESULT CMultiplicationDivision::MultiplicationOfTwoIntegers(int Num1, int Num2, int* pMultiplication)
{
	//code
	*pMultiplication = Num1 * Num2;
	return(S_OK);
}

//Imlementation of IDivision's methods
HRESULT CMultiplicationDivision::DivisionOfTwoIntegers(int Num1, int Num2, int* pDivision)
{
	//code
	*pDivision = Num1 / Num2;
	return(S_OK);
}

//Implemntation of CMultiplicationDivisionClassFactory's constructor method
CMultiplicationDivisionClassFactory::CMultiplicationDivisionClassFactory(void)
{
	//code
	m_cRef = 1; //hardcoded initialization to anticipate passible failuer of QueryInterface()
}

//Implemntation of CMultiplicationDivisionClassFactory's destructor method
CMultiplicationDivisionClassFactory::~CMultiplicationDivisionClassFactory(void)
{
	//code
}

//Implementation of CMultiplicationDivisionClassFactory's IClassFactory's IUnknown's Methods
HRESULT CMultiplicationDivisionClassFactory::QueryInterface(REFIID riid, void** ppv)
{
	//code
	if (riid == IID_IUnknown)
		*ppv = static_cast<IClassFactory*>(this);
	else if (riid == IID_IClassFactory)
		*ppv = static_cast<IClassFactory*>(this);
	else
	{
		*ppv = NULL;
		return(E_NOINTERFACE);
	}
	reinterpret_cast<IUnknown*>(*ppv)->AddRef();
	return(S_OK);
}

ULONG CMultiplicationDivisionClassFactory::AddRef(void)
{
	//code
	InterlockedIncrement(&m_cRef);
	return(m_cRef);
}

ULONG CMultiplicationDivisionClassFactory::Release(void)
{
	//code
	InterlockedDecrement(&m_cRef);
	if (m_cRef == 0)
	{
		delete(this);
		return(0);
	}
	return(m_cRef);
}

//Implementation of CMultiplicationDivisionClassFactory's IClassFactory's Methods
HRESULT CMultiplicationDivisionClassFactory::CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv)
{
	//veariable declarations
	CMultiplicationDivision *pCMultiplicationDivision = NULL;
	HRESULT hr;

	//code
	if ((pUnkOuter != NULL) && (riid != IID_IUnknown))
		return (CLASS_E_NOAGGREGATION);

	//create the instance of componant i.e of CMultiplicationDivision class
	pCMultiplicationDivision = new CMultiplicationDivision(pUnkOuter);
	if (pCMultiplicationDivision == NULL)
		return(E_OUTOFMEMORY);

	//get the vrequested interface
	hr = pCMultiplicationDivision->QueryInterface_NoAggregation(riid, ppv);
	pCMultiplicationDivision->Release_NoAggregation(); //anticipate possible failuer of QueryInterface()
	return(hr);
}

HRESULT CMultiplicationDivisionClassFactory::LockServer(BOOL fLock)
{
	//code
	if (fLock)
		InterlockedIncrement(&glNumberOfServerLocks);
	else
		InterlockedDecrement(&glNumberOfServerLocks);
	return(S_OK);
}

//Implementation of Expoted function from this Dll
HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
	//veriable declarations
	CMultiplicationDivisionClassFactory* pCMultiplicationDivisionClassFactory = NULL;
	HRESULT hr;

	//code
	if (rclsid != CLSID_MultiplicationDivision)
		return(CLASS_E_CLASSNOTAVAILABLE);

	//craete class factory
	pCMultiplicationDivisionClassFactory = new CMultiplicationDivisionClassFactory;
	if (pCMultiplicationDivisionClassFactory == NULL)
		return(E_OUTOFMEMORY);

	hr = pCMultiplicationDivisionClassFactory->QueryInterface(riid, ppv);
	pCMultiplicationDivisionClassFactory->Release();//anticipate possible failuer of QueryInterface()
	return(hr);
}

HRESULT __stdcall DllCanUnloadNow(void)
{
	//code
	if ((glNumberOfActiveComponants == 0) && (glNumberOfServerLocks == 0))
		return(S_OK);
	else
		return(S_FALSE);
}


