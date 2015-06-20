#include "stdafx.h"
#include "RibbonFramework.h"
#include "RibbonApplication.h"

IUIFramework* g_pRibbonFramework = nullptr;
IUIImageFromBitmap* g_pifbFactory = nullptr;

//
//  FUNCTION: InitializeRibbonFramework(HWND)
//
//  PURPOSE:  Initialize the Ribbon framework and bind a Ribbon to the application.
//
//  COMMENTS:
//
//    To get a Ribbon to display, the Ribbon framework must be initialized. 
//    This involves three important steps:
//      1) Instantiating the Ribbon framework object (CLSID_UIRibbonFramework).
//      2) Passing the host HWND and IUIApplication object to the framework.
//      3) Loading the binary markup compiled by UICC.exe.
//
//
bool InitializeRibbonFramework(HWND hWindowFrame)
{
    // Here we instantiate the Ribbon framework object.
    HRESULT hr = CoCreateInstance(CLSID_UIRibbonFramework, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&g_pRibbonFramework));
    if (FAILED(hr))
    {
        return false;
    }

    // Next, we create the application object (IUIApplication) and call the framework Initialize method, 
    // passing the application object and the host HWND that the Ribbon will attach itself to.
    CRibbonApplication *pRibbonApplication = nullptr;
    hr = CRibbonApplication::CreateInstance(&pRibbonApplication, hWindowFrame);
    if (FAILED(hr))
    {
        return false;
    }

    hr = g_pRibbonFramework->Initialize(hWindowFrame, pRibbonApplication);
    if (FAILED(hr))
    {
        pRibbonApplication->Release();
        return false;
    }

    // Finally, we load the binary markup.  This will initiate callbacks to the IUIApplication object 
    // that was provided to the framework earlier, allowing command handlers to be bound to individual
    // commands.
    hr = g_pRibbonFramework->LoadUI(GetModuleHandle(NULL), L"APPLICATION_RIBBON");
    if (FAILED(hr))
    {
        pRibbonApplication->Release();
        return false;
    }

	if (g_pifbFactory == nullptr)
	{
		hr = CoCreateInstance(CLSID_UIRibbonImageFromBitmapFactory, NULL, CLSCTX_ALL, IID_PPV_ARGS(&g_pifbFactory));
		if (FAILED(hr))
		{
			pRibbonApplication->Release();
			return false;
		}
	}

    pRibbonApplication->Release();

    return true;
}

//
//  FUNCTION: DestroyFramework()
//
//  PURPOSE:  Tears down the Ribbon framework.
//
//
void DestroyRibbonFramework()
{
    if (g_pRibbonFramework)
    {
        g_pRibbonFramework->Destroy();
        g_pRibbonFramework->Release();
        g_pRibbonFramework = NULL;
    }
}

//
//  FUNCTION: GetRibbonHeight()
//
//  PURPOSE:  Get the ribbon height.
//
//
UINT GetRibbonHeight()
{
    UINT ribbonHeight = 0U;
    if (g_pRibbonFramework)
    {
        IUIRibbon* pRibbon = nullptr;    
        HRESULT hr = g_pRibbonFramework->GetView(0, IID_PPV_ARGS(&pRibbon));

		if (SUCCEEDED(hr))
		{
			hr = pRibbon->GetHeight(&ribbonHeight);
			
			if (FAILED(hr))
			{
				ribbonHeight = 0;
			}
			pRibbon->Release();
		}
    }
    return ribbonHeight;
}