// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved.

#pragma once
#include <uiribbon.h>

//
//  CLASS: CRibbonHandler : IUICommandHandler
//
//  PURPOSE: Implements interface IUICommandHandler. 
//
//  COMMENTS:
//
//    This is the command handler for the buttons in the Ribbon
//    IUICommandHandler should be returned by the application during command creation.
//
class CRibbonHandler
    : public IUICommandHandler // Command handlers must implement IUICommandHandler.
{
public:
	HRESULT CreateUIImageFromBitmapResource(LPCTSTR pszResource,__out IUIImage **ppimg);
    STDMETHOD(Execute)(UINT nCmdID,
                       UI_EXECUTIONVERB verb, 
                       __in_opt const PROPERTYKEY* key,
                       __in_opt const PROPVARIANT* ppropvarValue,
                       __in_opt IUISimplePropertySet* pCommandExecutionProperties);

    STDMETHOD(UpdateProperty)(UINT nCmdID,
                              __in REFPROPERTYKEY key,
                              __in_opt const PROPVARIANT* ppropvarCurrentValue,
                              __out PROPVARIANT* ppropvarNewValue);
    
    static HRESULT CreateInstance(__deref_out CRibbonHandler **ppHandler);

    STDMETHOD_(ULONG, AddRef)();
    STDMETHOD_(ULONG, Release)();
    STDMETHOD(QueryInterface)(REFIID iid, void **ppv);

private:
	IUIImage *_IUIImageRunLarge = nullptr;
	IUIImage *_IUIImageRunSmall = nullptr;
	IUIImage *_IUIImageStopLarge = nullptr;
	IUIImage *_IUIImageStopSmall = nullptr;
	IUIImage *_IUIImageMinLarge = nullptr;
	IUIImage *_IUIImageMinSmall = nullptr;
	IUIImage *_IUIImageMaxLarge = nullptr;
	IUIImage *_IUIImageMaxSmall = nullptr;
	IUIImage *_IUIImageMeanLarge = nullptr;
	IUIImage *_IUIImageMeanSmall = nullptr;
	IUIImage *_IUIImageVarianceLarge = nullptr;
	IUIImage *_IUIImageVarianceSmall = nullptr;

    LONG m_cRef = 1L;
};