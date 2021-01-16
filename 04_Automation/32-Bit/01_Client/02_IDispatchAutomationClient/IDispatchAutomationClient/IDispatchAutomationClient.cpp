#include<windows.h>
#include<stdio.h>
#include"AutomationServerWithRegFile.h"
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

//WinMain()
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpszCmdLine, int iCmdShow)
{
    WNDCLASSEX wndclass;
    HWND hwnd;
    MSG msg;
    TCHAR szAppName[] = TEXT("Client");
  
    wndclass.cbSize = sizeof(WNDCLASSEX);
    wndclass.style = CS_HREDRAW | CS_VREDRAW;
    wndclass.cbClsExtra = 0;
    wndclass.cbWndExtra = 0;
    wndclass.lpfnWndProc = WndProc;
    wndclass.hInstance = hInstance;
    wndclass.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wndclass.hCursor = LoadCursor(NULL, IDC_ARROW);
    wndclass.hbrBackground = (HBRUSH)GetStockObject(BLACK_BRUSH);
    wndclass.lpszClassName = szAppName;
    wndclass.lpszMenuName = NULL;
    wndclass.hIconSm = LoadIcon(NULL, IDI_APPLICATION);

    RegisterClassEx(&wndclass);

    hwnd = CreateWindow(szAppName,
        TEXT("Client Of COM Dll Server"),
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        NULL,
        NULL,
        hInstance,
        NULL);

    ShowWindow(hwnd, iCmdShow);
    UpdateWindow(hwnd);

    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
     return((int)msg.wParam);
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
    //function declarartion
    void ComErrorDescriptionString(HWND, HRESULT);

    //veriable declartions
    IDispatch* pIDispatch = NULL;
    HRESULT hr;
    DISPID dispid;
    WCHAR  szFunctionOne[] = L"SumOfTwoIntegers"; 
    OLECHAR* szFunctionNameOne = szFunctionOne;
    
    WCHAR  szFunctionTwo[] = L"SubtractionOfTwoIntegers"; 
    OLECHAR* szFunctionNameTwo = szFunctionTwo;

    VARIANT vArg[2], vRet;
    DISPPARAMS param = { vArg, 0, 2, NULL};
    int n1,n2;
    TCHAR str[255];

    //code
    switch (iMsg)
    {
    case WM_CREATE:
        //Initialized Com Library
        hr = CoInitialize(NULL);
        if (FAILED(hr))
        {
            MessageBox(NULL, TEXT("COM Library Can Not Be Initialized.\nProgram Will Now Exit."),
                TEXT("Program Error"), MB_OK);
            exit(0);
        }

        hr = CoCreateInstance(CLSID_MyMath,
            NULL,
            CLSCTX_INPROC_SERVER,
            IID_IDispatch,
            (void**)&pIDispatch
        );
        if (FAILED(hr))
        {
            MessageBox(hwnd, TEXT("Componat Can Not Be Obtained."), TEXT("Error"), MB_OK | MB_ICONERROR | MB_TOPMOST);
            DestroyWindow(hwnd);
            exit(0);
        }

        //initialized arguments hardcoded
        //Commaon code for both IMyMath->SumOfTwoIntegers() and IMyMath->SubtractionOftwoInteges()
        n1 = 21;
        n2 = 11;

        VariantInit(vArg);
        vArg[0].vt = VT_INT;
        vArg[0].intVal = n2;
        vArg[1].vt = VT_INT;
        vArg[1].intVal = n1;

        param.cArgs = 2;
        param.cNamedArgs = 0;
        param.rgdispidNamedArgs = NULL;

        //reverse order of parameter
        param.rgvarg = vArg;

        //return value
        VariantInit(&vRet);

        //*********code for IMyMath->SumOfTwoIntegers()*********
        hr = pIDispatch->GetIDsOfNames(IID_NULL,
             &szFunctionNameOne,
             1,
             GetUserDefaultLCID(),
             &dispid);
       if (FAILED(hr))
        {
           ComErrorDescriptionString(hwnd, hr);
           MessageBox(hwnd, TEXT("Canot get Id for SumOfTwoIntegers()"), TEXT("Error"), MB_OK | MB_ICONERROR | MB_TOPMOST);
           pIDispatch->Release();
           DestroyWindow(hwnd);
        }

       hr = pIDispatch->Invoke(dispid,
           IID_NULL,
           GetUserDefaultLCID(),
           DISPATCH_METHOD,
           &param,
           &vRet,
           NULL,
           NULL);
       if (FAILED(hr))
       {
           ComErrorDescriptionString(hwnd, hr);
           MessageBox(hwnd, TEXT("Canot Invoke Function"), TEXT("Error"), MB_OK | MB_ICONERROR | MB_TOPMOST);
           pIDispatch->Release();
           DestroyWindow(hwnd);
       }
       else
       {
           //display result
           wsprintf(str, TEXT("Sum of %d and %d is: %d"), n1,n2, vRet.lVal);
           MessageBox(hwnd, str, TEXT("SumOfTwoIntegers"), MB_OK);
       }

       //*********code for  IMyMath->SubtractionOfTwoIntegers()*********
       hr = pIDispatch->GetIDsOfNames(IID_NULL,
           &szFunctionNameTwo,
           1,
           GetUserDefaultLCID(),
           &dispid);
       if (FAILED(hr))
       {
           ComErrorDescriptionString(hwnd, hr);
           MessageBox(hwnd, TEXT("Canot get Id for SubtractionOfTwoIntegers()"), TEXT("Error"), MB_OK | MB_ICONERROR | MB_TOPMOST);
           pIDispatch->Release();
           DestroyWindow(hwnd);
       }

       hr = pIDispatch->Invoke(dispid,
           IID_NULL,
           GetUserDefaultLCID(),
           DISPATCH_METHOD,
           &param,
           &vRet,
           NULL,
           NULL);
       if (FAILED(hr))
       {
           ComErrorDescriptionString(hwnd, hr);
           MessageBox(hwnd, TEXT("Canot Invoke Function"), TEXT("Error"), MB_OK | MB_ICONERROR | MB_TOPMOST);
           pIDispatch->Release();
           DestroyWindow(hwnd);
       }
       else
       {
           //display result
           wsprintf(str, TEXT("Subtraction of %d and %d is: %d"), n1, n2, vRet.lVal);
           MessageBox(hwnd, str, TEXT("SubtractionOfTwoIntegers"), MB_OK);
       }

       //clean-up
       VariantClear(vArg);
       VariantClear(&vRet);
       pIDispatch->Release();
       pIDispatch = NULL;
       DestroyWindow(hwnd);
       break;

    case WM_DESTROY:
        CoUninitialize();
        PostQuitMessage(0);
        break;

    }
    return (DefWindowProc(hwnd, iMsg, wParam, lParam));
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
