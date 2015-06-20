// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved.

#pragma once

#define MAX_RESOURCE_LENGTH     256
#include <uiribbon.h>

// The implementation of IUISimplePropertySet. This handles all of the properties used for the 
// ItemsSource and Categories PKEYs and provides functions to set only the properties required 
// for each type of gallery contents.
class CRibbonSimplePropertySet : public IUISimplePropertySet
{
public:

    static HRESULT CreateInstance(__deref_out CRibbonSimplePropertySet **ppPropertySet);

    void InitializeCommandProperties(int categoryId, int commandId, UI_COMMANDTYPE commandType);

    void InitializeItemProperties(IUIImage *image, __in PCWSTR label, int categoryId);

    void InitializeCategoryProperties(__in PCWSTR label, int categoryId);

    STDMETHOD(GetValue)(__in REFPROPERTYKEY key, __out PROPVARIANT *ppropvar);

    STDMETHOD_(ULONG, AddRef)();

    STDMETHOD_(ULONG, Release)();

    STDMETHOD(QueryInterface)(REFIID iid, void **ppv);

private:

    ~CRibbonSimplePropertySet()
    {
		SafeRelease(&m_pimgItem);
    }

	// Used for items and categories
	WCHAR m_wszLabel[MAX_RESOURCE_LENGTH];
	// Used for items, categories, and commands
    int m_categoryId = UI_COLLECTION_INVALIDINDEX;
	// Used for items only
    IUIImage* m_pimgItem = nullptr;
	// Used for commands only
    int m_commandId = -1;
	// Used for commands only
    UI_COMMANDTYPE m_commandType = UI_COMMANDTYPE_UNKNOWN;

    LONG m_cRef = 1L;
};