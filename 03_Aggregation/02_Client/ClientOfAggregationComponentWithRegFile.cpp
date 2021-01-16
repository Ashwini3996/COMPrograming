#include<windows.h>
#include"HeaderForClientOfAggregationComponentWithRegFile.h"

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

//global veriable declarations
ISum* pISum = NULL;
ISubtract* pISubtract = NULL;
IMultiplication* pIMultiplication = NULL;
IDivision* pIDivision = NULL;

//WinMain()
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpszCmdLine, int iCmdShow)
{
    WNDCLASSEX wndclass;
    HWND hwnd;
    MSG msg;
    TCHAR szAppName[] = TEXT("ComClient");
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

    //veriable declartions
    HRESULT hr;
    int iNum1, iNum2, iSum, iSubtraction, iMultiplication, iDivision;
    TCHAR str[255];

    //code
    switch (iMsg)
    {
    case WM_CREATE:
        hr = CoCreateInstance(CLSID_SumSubtract,
            NULL,
            CLSCTX_INPROC_SERVER,
            IID_ISum,
            (void**)&pISum
        );
        if (FAILED(hr))
        {
            MessageBox(hwnd, TEXT("ISum Interface Can Not Be Obtained."), TEXT("Error"), MB_OK);
            DestroyWindow(hwnd);
        }

        //initialized arguments hardcoded
        iNum1 = 50;
        iNum2 = 25;

        //call SumOfTwoIntegers() of ISum to get the sum
        pISum->SumOfTwoIntegers(iNum1, iNum2, &iSum);
        wsprintf(str, TEXT("Sum of %d and %d is: %d"), iNum1, iNum2, iSum);
        MessageBox(hwnd, str, TEXT("Result"), MB_OK);

        //call QueryInterface() on ISum to get ISubtracts pointer
        hr = pISum->QueryInterface(IID_ISubtract, (void**)&pISubtract);
        if (FAILED(hr))
        {
            MessageBox(hwnd, TEXT("ISubtact Interface Can Not Be Obtained."), TEXT("Error"), MB_OK);
            DestroyWindow(hwnd);
        }

        //as ISum is now not nedded onwords, relesed it
        pISum->Release();
        pISum = NULL; //Make relesed interface NULL

        //again initialized arguments hardcoded
        iNum1 = 50;
        iNum2 = 25;

        //again call SubtractionOfTwoIntegers() of ISubtract to get the subtraction
        pISubtract->SubtractionOfTwoIntegers(iNum1, iNum2, &iSubtraction);

        //display result
        wsprintf(str, TEXT("Subtraction of %d and %d is: %d"), iNum1, iNum2, iSubtraction);
        MessageBox(hwnd, str, TEXT("Result"), MB_OK);

        //call QueryInterface() on ISubtract's to get IMultiplication pointer
        hr = pISubtract->QueryInterface(IID_IMultiplication, (void**)&pIMultiplication);
        if (FAILED(hr))
        {
            MessageBox(hwnd, TEXT("IMultiplication Interface Can Not Be Obtained."), TEXT("Error"), MB_OK);
            DestroyWindow(hwnd);
        }

        //as ISubtract is now not nedded onwords, relesed it
        pISubtract->Release();
        pISubtract = NULL; //Make relesed interface NULL

        iNum1 = 50;
        iNum2 = 25;

        pIMultiplication->MultiplicationOfTwoIntegers(iNum1, iNum2, &iMultiplication);

        //display result
        wsprintf(str, TEXT("Multiplication of %d and %d is: %d"), iNum1, iNum2, iMultiplication);
        MessageBox(hwnd, str, TEXT("Result"), MB_OK);


        //call QueryInterface() on IMultiplication's to get IDivision pointer
        hr = pIMultiplication->QueryInterface(IID_IDivision, (void**)&pIDivision);
        if (FAILED(hr))
        {
            MessageBox(hwnd, TEXT("IDivision Interface Can Not Be Obtained."), TEXT("Error"), MB_OK);
            DestroyWindow(hwnd);
        }

        //as IMultiplication is now not nedded onwords, relesed it
        pIMultiplication->Release();
        pIMultiplication = NULL; //Make relesed interface NULL

        iNum1 = 50;
        iNum2 = 25;
        pIDivision->DivisionOfTwoIntegers(iNum1, iNum2, &iDivision);

        //display result
        wsprintf(str, TEXT("Division of %d and %d is: %d"), iNum1, iNum2, iDivision);
        MessageBox(hwnd, str, TEXT("Result"), MB_OK);

        pIDivision->Release();
        pIDivision = NULL;

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


void SafeInterfaceRelease(void)//at a time remove all interfaces
{
    //code
    if (pISum)
    {
        pISum->Release();
        pISum = NULL;
    }

    if (pISubtract)
    {
        pISubtract->Release();
        pISubtract = NULL;
    }

    if (pIMultiplication)
    {
        pIMultiplication->Release();
        pIMultiplication = NULL;
    }

    if (pIDivision)
    {
        pIDivision->Release();
        pIDivision = NULL;
    }
}

