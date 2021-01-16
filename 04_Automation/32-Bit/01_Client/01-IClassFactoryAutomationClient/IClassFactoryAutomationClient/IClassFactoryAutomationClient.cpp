#include<windows.h>
#include<stdio.h>
#include"AutomationServerWithRegFile.h"
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

//global veriable declarations
IMyMath* pIMyMath = NULL;

//WinMain()
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpszCmdLine, int iCmdShow)
{
    WNDCLASSEX wndclass;
    HWND hwnd;
    MSG msg;
    TCHAR szAppName[] = TEXT("MyComExe");
    HRESULT hr;

    //code
    //COM Initialization
    hr = CoInitialize(NULL);
    if (FAILED(hr))
    {
        MessageBox(NULL, TEXT("COM Library Can Not Be Initialized.\nProgram Will Now Exit."),
            TEXT("Program Error"), MB_OK);
        exit(0);
    }

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
    //COM Uninitialization
    CoUninitialize();
    return((int)msg.wParam);
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
    //function declarartion
    void SafeInterfaceRelease(void);
    void ComErrorDescriptionString(HWND, HRESULT);

    //veriable declartions
    HRESULT hr;
    int iNum1, iNum2, iSum, iSubtract;
    TCHAR str[255];

    //code
    switch (iMsg)
    {
    case WM_CREATE:
        hr = CoCreateInstance(CLSID_MyMath,
            NULL,
            CLSCTX_INPROC_SERVER,
            IID_IMyMath,
            (void**)&pIMyMath
        );
        if (FAILED(hr))
        {
            ComErrorDescriptionString(hwnd, hr);
            DestroyWindow(hwnd);
        }

        //initialized arguments hardcoded
        iNum1 = 21;
        iNum2 = 11;

        //call SumOfTwoIntegers() of IMyMath to get the sum
        pIMyMath->SumOfTwoIntegers(iNum1, iNum2, &iSum);
        wsprintf(str, TEXT("Sum of %d and %d is: %d"), iNum1, iNum2, iSum);
        MessageBox(hwnd, str, TEXT("Result"), MB_OK);

        //call SumOfTwoIntegers() of IMyMath to get the sum
        pIMyMath->SubtractionOfTwoIntegers(iNum1, iNum2, &iSubtract);
        wsprintf(str, TEXT("Subtraction of %d and %d is: %d"), iNum1, iNum2, iSubtract);
        MessageBox(hwnd, str, TEXT("Result"), MB_OK);

        pIMyMath->Release();
        pIMyMath = NULL; //Make relesed interface NULL
        
        //exit application
        DestroyWindow(hwnd);
        break;

    case WM_DESTROY:
        SafeInterfaceRelease();
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

void SafeInterfaceRelease(void)
{
    if (pIMyMath)
    {
        pIMyMath->Release();
        pIMyMath = NULL;
    }
}