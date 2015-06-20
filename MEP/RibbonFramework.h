//
// RibbonFrameword.h/cpp implements the initialization and tear down of ribbon framework.
//

#pragma once
#include <uiribbon.h>

extern IUIFramework* g_pRibbonFramework; // Reference to the ribbon framework.
extern IUIImageFromBitmap* g_pifbFactory; // Reference to the ribbon framework.

// Methods to setup and tear down the framework.
bool InitializeRibbonFramework(HWND hWindowFrame);
void DestroyRibbonFramework();

UINT GetRibbonHeight();