#include<windows.h>
#include"AggregationOuterComponentWithRegFile.h"
#include"AggregationInnerComponentWithRegFile.h"

//class declaration

class CSumSubtract :public ISum, ISubtract
{
private:
	long m_cRef;
	IUnknown* m_pIUnknownInner;
	//IMultiplication* m_pIMultiplication;
	//IDivision* m_pIDivision;

public:
	//constructor method declartion
	CSumSubtract(void);

	//destructor method declartion
	~CSumSubtract(void);

	//IUnknown specific method declartion(inherited)
	HRESULT __stdcall QueryInterface(REFIID, void**);
	ULONG __stdcall   AddRef(void);
	ULONG __stdcall   Release(void);

	//ISum specific method declartion(inherited)
	HRESULT __stdcall SumOfTwoIntegers(int, int, int*);

	//ISubtract specific method declartion(inherited)
	HRESULT __stdcall SubtractionOfTwoIntegers(int, int, int*);

	//custom method for inner component creation
	HRESULT __stdcall InitializeInnerComponent(void);
};

class CSumSubtractClassFactory :public IClassFactory
{
private:
	long m_cRef;

public:
	//constructor method declartion
	CSumSubtractClassFactory(void);

	//destructor method declartion
	~CSumSubtractClassFactory(void);

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
CSumSubtract::CSumSubtract(void)
{
	//Initialized private data members
	m_pIUnknownInner = NULL;
	//m_pIMultiplication = NULL;
	//m_pIDivision = NULL;
	m_cRef = 1; //hardcoded initialization to anticipate passible failuer of QueryInterface()
	InterlockedIncrement(&glNumberOfActiveComponants);//Increment global counter
}

//Implemetation of CSumSubtract's Destrucyor method
CSumSubtract::~CSumSubtract(void)
{
	InterlockedDecrement(&glNumberOfActiveComponants);//Decrement global counter
   /*if (m_pIMultiplication)
	{
		m_pIMultiplication->Release();
		m_pIMultiplication = NULL;
	}

	if (m_pIDivision)
	{
		m_pIDivision->Release();
		m_pIDivision = NULL;
	}*/


	if (m_pIUnknownInner)
	{
		m_pIUnknownInner->Release();
		m_pIUnknownInner = NULL;
	}
}

//Implemetation of CSumSubtract's IUnknown method
HRESULT CSumSubtract::QueryInterface(REFIID riid, void** ppv)
{
	//code
	if (riid == IID_IUnknown)
		*ppv = static_cast<ISum*>(this);
	else if (riid == IID_ISum)
		*ppv = static_cast<ISum*>(this);
	else if (riid == IID_ISubtract)
		*ppv = static_cast<ISubtract*>(this);
	else if (riid == IID_IMultiplication)
		 return(m_pIUnknownInner-> QueryInterface(riid, ppv));
	else if (riid == IID_IDivision)
		return(m_pIUnknownInner->QueryInterface(riid, ppv));
	else
	{
		*ppv = NULL;
		return(E_NOINTERFACE);
	}
	reinterpret_cast<IUnknown*>(*ppv)->AddRef();
	return(S_OK);
}

ULONG CSumSubtract::AddRef(void)
{
	//code
	InterlockedIncrement(&m_cRef);
	return(m_cRef);
}

ULONG CSumSubtract::Release(void)
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
HRESULT CSumSubtract::SumOfTwoIntegers(int Num1, int Num2, int* pSum)
{
	//code
	*pSum = Num1 + Num2;
	return(S_OK);
}

//Imlementation of ISubtract's methods
HRESULT CSumSubtract::SubtractionOfTwoIntegers(int Num1, int Num2, int* pSubtract)
{
	//code
	*pSubtract = Num1 - Num2;
	return(S_OK);
}

HRESULT CSumSubtract::InitializeInnerComponent(void)
{
	//veriable declaration
	HRESULT hr;

	//code
	hr = CoCreateInstance(CLSID_MultiplicationDivision, 
		            reinterpret_cast<IUnknown *>(this),
		                          CLSCTX_INPROC_SERVER, 
		                                  IID_IUnknown,
		                     (void**)&m_pIUnknownInner
	                     );

	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("IUnknown Interface Can Not Be Obtained From Inner Commponet."),
			TEXT("Error"), MB_OK);
		return(E_FAIL);
	}

	return(S_OK);
}

//Implemntation of CSumSubtractClassFactory's constructor method
CSumSubtractClassFactory::CSumSubtractClassFactory(void)
{
	//code
	m_cRef = 1; //hardcoded initialization to anticipate passible failuer of QueryInterface()
}

//Implemntation of CSumSubtractClassFactory's destructor method
CSumSubtractClassFactory::~CSumSubtractClassFactory(void)
{
	//code
}

//Implementation of CSumSubtractClassFactory's IClassFactory's IUnknown's Methods
HRESULT CSumSubtractClassFactory::QueryInterface(REFIID riid, void** ppv)
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

ULONG CSumSubtractClassFactory::AddRef(void)
{
	//code
	InterlockedIncrement(&m_cRef);
	return(m_cRef);
}

ULONG CSumSubtractClassFactory::Release(void)
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
HRESULT CSumSubtractClassFactory::CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv)
{
	//veariable declarations
	CSumSubtract* pCSumSubtract = NULL;
	HRESULT hr;

	//code
	if (pUnkOuter != NULL)
		return (CLASS_E_NOAGGREGATION);

	//create the instance of componant i.e of CSumSubtract class
	pCSumSubtract = new CSumSubtract;
	if (pCSumSubtract == NULL)
		return(E_OUTOFMEMORY);

	//initialized inner component
	hr = pCSumSubtract->InitializeInnerComponent();
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("Failed To Initialized Inner Commponet."), TEXT("Error"), MB_OK);
		pCSumSubtract->Release();
		return(E_FAIL);
	}

	//get the vrequested interface
	hr = pCSumSubtract->QueryInterface(riid, ppv);
	pCSumSubtract->Release(); //anticipate possible failuer of QueryInterface()
	return(hr);
}

HRESULT CSumSubtractClassFactory::LockServer(BOOL fLock)
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
	CSumSubtractClassFactory* pCSumSubtractClassFactory = NULL;
	HRESULT hr;

	//code
	if (rclsid != CLSID_SumSubtract)
		return(CLASS_E_CLASSNOTAVAILABLE);

	//craete class factory
	pCSumSubtractClassFactory = new CSumSubtractClassFactory;
	if (pCSumSubtractClassFactory == NULL)
		return(E_OUTOFMEMORY);

	hr = pCSumSubtractClassFactory->QueryInterface(riid, ppv);
	pCSumSubtractClassFactory->Release();//anticipate possible failuer of QueryInterface()
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


