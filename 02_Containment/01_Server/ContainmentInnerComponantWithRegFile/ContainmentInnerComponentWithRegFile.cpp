#include<windows.h>
#include"ContainmentInnerComponentWithRegFile.h"

//class declaration

class CMultiplicationDivision :public IMultiplication, IDivision
{
private:
	long m_cRef;
	
public:
	//constructor method declartion
	CMultiplicationDivision(void);

	//destructor method declartion
	~CMultiplicationDivision(void);

	//IUnknown specific method declartion(inherited)
	HRESULT __stdcall QueryInterface(REFIID, void**);
	ULONG __stdcall   AddRef(void);
	ULONG __stdcall   Release(void);
	
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

//Implemetation of CSumSubtract's Constructor method
CMultiplicationDivision::CMultiplicationDivision(void)
{
	 m_cRef = 1; //hardcoded initialization to anticipate passible failuer of QueryInterface()
	InterlockedIncrement(&glNumberOfActiveComponants);//Increment global counter
}

//Implemetation of CSumSubtract's Destrucyor method
CMultiplicationDivision::~CMultiplicationDivision(void)
{
	InterlockedDecrement(&glNumberOfActiveComponants);//Decrement global counter
}

//Implemetation of CSumSubtract's IUnknown method
HRESULT CMultiplicationDivision::QueryInterface(REFIID riid, void** ppv)
{
	//code
	if (riid == IID_IUnknown)
		*ppv = static_cast<IMultiplication *>(this);
	else if (riid == IID_IMultiplication)
		*ppv = static_cast<IMultiplication *>(this);
	else if (riid == IID_IDivision)
		*ppv = static_cast<IDivision *>(this);
	else
	{
		*ppv = NULL;
		return(E_NOINTERFACE);
	}
	reinterpret_cast<IUnknown*>(*ppv)->AddRef();
	return(S_OK);
}

ULONG CMultiplicationDivision::AddRef(void)
{
	//code
	InterlockedIncrement(&m_cRef);
	return(m_cRef);
}

ULONG CMultiplicationDivision::Release(void)
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

//Imlementation of ISum's methods
HRESULT CMultiplicationDivision::MultiplicationOfTwoIntegers(int Num1, int Num2, int* pMultiplication)
{
	//code
	*pMultiplication = Num1 * Num2;
	return(S_OK);
}

//Imlementation of ISubtract's methods
HRESULT CMultiplicationDivision::DivisionOfTwoIntegers(int Num1, int Num2, int* pDivision)
{
	//code
	*pDivision = Num1 / Num2;
	return(S_OK);
}

//Implemntation of CSumSubtractClassFactory's constructor method
CMultiplicationDivisionClassFactory::CMultiplicationDivisionClassFactory(void)
{
	//code
	m_cRef = 1; //hardcoded initialization to anticipate passible failuer of QueryInterface()
}

//Implemntation of CSumSubtractClassFactory's destructor method
CMultiplicationDivisionClassFactory::~CMultiplicationDivisionClassFactory(void)
{
	//code
}

//Implementation of CSumSubtractClassFactory's IClassFactory's IUnknown's Methods
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

//Implementation of CSumSubtractClassFactory's IClassFactory's Methods
HRESULT CMultiplicationDivisionClassFactory::CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv)
{
	//veariable declarations
	CMultiplicationDivision * pCMultiplicationDivision = NULL;
	HRESULT hr;

	//code
	if (pUnkOuter != NULL)
		return (CLASS_E_NOAGGREGATION);

	//create the instance of componant i.e of CSumSubtract class
	pCMultiplicationDivision = new CMultiplicationDivision;
	if (pCMultiplicationDivision == NULL)
		return(E_OUTOFMEMORY);

	//get the vrequested interface
	hr = pCMultiplicationDivision->QueryInterface(riid, ppv);
	pCMultiplicationDivision->Release(); //anticipate possible failuer of QueryInterface()
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
	CMultiplicationDivisionClassFactory * pCMultiplicationDivisionClassFactory = NULL;
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


