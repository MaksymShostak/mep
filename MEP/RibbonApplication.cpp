#include "stdafx.h"
#include "RibbonApplication.h"
#include "RibbonHandler.h"

extern HWND hWnd;

//
//  FUNCTION: OnViewChanged(UINT, UI_VIEWTYPE, IUnknown*, UI_VIEWVERB, INT)
//
//  PURPOSE: Called when the state of a View (Ribbon is a view) changes, for example, created, destroyed, or resized.
//
//
STDMETHODIMP CRibbonApplication::OnViewChanged(UINT32 /*nViewID*/, __in UI_VIEWTYPE typeID,	__in IUnknown* /*pView*/, UI_VIEWVERB verb, INT32 /*uReasonCode*/)
{
    HRESULT hr = E_FAIL;

	// Checks to see if the view that was changed was a Ribbon view.
    if (UI_VIEWTYPE_RIBBON == typeID)
    {
		switch (verb)
		{
		 // The view was newly created.
         case UI_VIEWVERB_CREATE:
			 {
                hr = S_OK;
			 }
			 break;
		
		// The view has been resized.
		case UI_VIEWVERB_SIZE:
			{
				RECT recthWnd = {0};

				GetClientRect(hWnd, &recthWnd);
				LPARAM lParam = MAKELPARAM(recthWnd.right-recthWnd.left, recthWnd.bottom-recthWnd.top);

				SendMessageW(hWnd, WM_SIZE, NULL, lParam);

				/*IUIRibbon* pRibbon = NULL;
                UINT uRibbonHeight;
                // Call to the framework to determine the height of the Ribbon.
                hr = pView->QueryInterface(IID_PPV_ARGS(&pRibbon));
                if (SUCCEEDED(hr))
                {
                    hr = pRibbon->GetHeight(&uRibbonHeight);
                    pRibbon->Release();
                }*/

				hr = S_OK;
			}
			break;

		// The view was destroyed.
        case UI_VIEWVERB_DESTROY:
			{
                hr = S_OK;
            }
			break;
		}
	}
    return hr;
}

//
//  FUNCTION: OnCreateUICommand(UINT, UI_COMMANDTYPE, IUICommandHandler)
//
//  PURPOSE: Called by the Ribbon framework for each command specified in markup, to allow
//           the host application to bind a command handler to that command.
//
//  COMMENTS:
//
//    In this Gallery sample, there is one handler for each gallery, and one for all of
//    the buttons in the Size and Color gallery.
//
//
STDMETHODIMP CRibbonApplication::OnCreateUICommand(UINT32 /*nCmdID*/, __in UI_COMMANDTYPE /*typeID*/, __deref_out IUICommandHandler** ppCommandHandler)
{
	CRibbonHandler* pRibbonHandler = nullptr;

	HRESULT hr = CRibbonHandler::CreateInstance(&pRibbonHandler);

	if (SUCCEEDED(hr))
	{
		hr = pRibbonHandler->QueryInterface(IID_PPV_ARGS(ppCommandHandler));

		SafeRelease(&pRibbonHandler);
	}

    return hr;
}

//
//  FUNCTION: OnDestroyUICommand(UINT, UI_COMMANDTYPE, IUICommandHandler*)
//
//  PURPOSE: Called by the Ribbon framework for each command at the time of ribbon destruction.
//
//
STDMETHODIMP CRibbonApplication::OnDestroyUICommand(UINT32 /*commandId*/, __in UI_COMMANDTYPE /*typeID*/, __in_opt IUICommandHandler* /*pCommandHandler*/)
{
    return E_NOTIMPL;
}

HRESULT CRibbonApplication::CreateInstance(__deref_out CRibbonApplication **ppRibbonApplication, HWND hwnd)
{
    if (!ppRibbonApplication)
    {
        return E_POINTER;
    }

    if (!hwnd)
    {
        return E_INVALIDARG;
    }

    *ppRibbonApplication = nullptr;

    HRESULT hr = S_OK;

    CRibbonApplication* pRibbonApplication = new CRibbonApplication();

    if (pRibbonApplication)
    {
        pRibbonApplication->m_hwnd = hwnd;
        *ppRibbonApplication = pRibbonApplication;
        
    }
    else
    {
        hr = E_OUTOFMEMORY;
    }

    return hr;
}

// IUnknown methods.
STDMETHODIMP_(ULONG) CRibbonApplication::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CRibbonApplication::Release()
{
    LONG cRef = InterlockedDecrement(&m_cRef);
    if (cRef == 0)
    {
        delete this;
    }

    return cRef;
}

STDMETHODIMP CRibbonApplication::QueryInterface(REFIID iid, void** ppv)
{
    if (!ppv)
    {
        return E_POINTER;
    }

    if (iid == __uuidof(IUnknown))
    {
        *ppv = static_cast<IUnknown*>(this);
    }
    else if (iid == __uuidof(IUIApplication))
    {
        *ppv = static_cast<IUIApplication*>(this);
    }
    else 
    {
        *ppv = NULL;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}