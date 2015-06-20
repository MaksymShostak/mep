// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved.

#include "stdafx.h"
#include "RibbonHandler.h"
#include "RibbonSimplePropertySet.h"
#include "Resource.h"
#include "ribbonres.h"
#include "RibbonFramework.h"
#include <uiribbonpropertyhelpers.h>

extern VARIABLES g_Variables;
extern HWND hWnd;

HRESULT CRibbonHandler::CreateUIImageFromBitmapResource(LPCTSTR pszResource,__out IUIImage **ppimg)
{
	HRESULT hr = E_FAIL;

	*ppimg = nullptr;

	// Load the bitmap from the resource file.
	HBITMAP hbm = static_cast<HBITMAP>(LoadImageW(GetModuleHandleW(nullptr), pszResource, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION));

	if (hbm)
	{
		// Use the factory implemented by the framework to produce an IUIImage.
		hr = g_pifbFactory->CreateImage(hbm, UI_OWNERSHIP_COPY, ppimg);

		(void)DeleteObject(hbm);
	}

	return hr;
}

STDMETHODIMP CRibbonHandler::Execute(UINT nCmdID,
                   UI_EXECUTIONVERB verb, 
                   __in_opt const PROPERTYKEY* key,
                   __in_opt const PROPVARIANT* ppropvarValue,
                   __in_opt IUISimplePropertySet* pCommandExecutionProperties)
{
    HRESULT hr = S_OK;
	
    if (verb == UI_EXECUTIONVERB_EXECUTE)
    {
		switch (nCmdID)
		{
		case ID_FILE_NEW:
		case ID_FILE_OPEN:
		case ID_FILE_OPEN_ENTROPYGRID:
		case ID_FILE_OPEN_LIFEGRID:
		case ID_FILE_SAVE:
		case ID_FILE_SAVE_ENTROPYGRID:
		case ID_FILE_SAVE_LIFEGRID:
		case ID_FILE_INFO:
		case ID_FILE_EXIT:
		case ID_HELP:
		case cmdHome_Options_Visualisation:
		case cmdHome_Options_OutputToFile:
		case cmdHome_Options_Paint:
		case cmdHome_Commands_Run:
		case cmdHome_Commands_NextStep:
		case cmdHome_Commands_Clear:
		case cmdHome_Commands_ClearEntropyGrid:
		case cmdHome_Commands_ClearLifeGrid:
		case cmdHome_Ruleset_PeriodicBoundaryConditions:
		case cmdHome_EntropyGrid_Initialise:
		case cmdHome_LifeGrid_Intelligence:
		case cmdHome_LifeGrid_Bias:
		case cmdHome_LifeGrid_Initialise:
			{
				PostMessageW(hWnd, WM_COMMAND, nCmdID, 0);
			}
			break;
		
		case cmdHome_Ruleset_RandomNumberFunction:
			{
				if (key && *key == UI_PKEY_SelectedItem)
				{
					UINT selected = ppropvarValue->uintVal;
					
					g_Variables.RandomNumberFunction = (RANDOM_NUMBER_FUNCTION)selected;
				}
			}
			break;
		
		case cmdHome_EntropyGrid_Initialisation:
			{
				if (key && *key == UI_PKEY_SelectedItem)
				{
					UINT selected = ppropvarValue->uintVal;
					
					g_Variables.Initialisation = (INITIALISATION)selected;

					/*hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_ALLPROPERTIES, NULL);
					if (FAILED(hr)) { return hr; }

					hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_ALLPROPERTIES, NULL);
					if (FAILED(hr)) { return hr; }*/

					hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_Label);
					if (FAILED(hr)) { return hr; }
					hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
					if (FAILED(hr)) { return hr; }
					hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_TooltipTitle);
					if (FAILED(hr)) { return hr; }
					hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_TooltipDescription);
					if (FAILED(hr)) { return hr; }
					hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_LargeImage);
					if (FAILED(hr)) { return hr; }
					hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_SmallImage);
					if (FAILED(hr)) { return hr; }

					hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_Label);
					if (FAILED(hr)) { return hr; }
					hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
					if (FAILED(hr)) { return hr; }
					hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_TooltipTitle);
					if (FAILED(hr)) { return hr; }
					hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_TooltipDescription);
					if (FAILED(hr)) { return hr; }
					hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_LargeImage);
					if (FAILED(hr)) { return hr; }
					hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_SmallImage);
					if (FAILED(hr)) { return hr; }
				}
			}
			break;
		case cmdHome_EntropyGrid_InitialisationVariable1:
			{
				if (key && *key == UI_PKEY_SelectedItem)
				{
					if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
					{
						UINT selected = ppropvarValue->uintVal;
						switch (selected)
						{
						case 0:
							g_Variables.UniformMin = 0;
							break;
						case UI_COLLECTION_INVALIDINDEX: // The new selection is a value that the user typed.
							if (pCommandExecutionProperties)
							{
								PROPVARIANT var;
								PropVariantInit(&var);

								hr = pCommandExecutionProperties->GetValue(UI_PKEY_Label, &var); // The text entered by the user is contained in the property set with the pkey UI_PKEY_Label.
								if (FAILED(hr)) { return hr; }

								BSTR bstr = var.bstrVal;
								ULONG newSize = 0UL;

								hr = VarUI4FromStr(bstr, GetUserDefaultLCID(), 0, &newSize);
                    
								if (FAILED(hr))
								{
									MessageBeep(MB_ICONERROR);
									MessageBox(NULL, L"VarUI4FromStr Error", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
									// Manually changing the text requires invalidating the StringValue property.
									hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
									break;
								}
								else if (newSize < 0 || newSize > (UINT)pow(2.0, (int)((g_Variables.BitsPerPixel - 8)/3)))
								{
									MessageBeep(MB_ICONERROR);
									MessageBox(NULL, L"0 <= x <= 256", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
									// Manually changing the text requires invalidating the StringValue property.
									hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
									break;
								}
								else if (newSize > g_Variables.UniformMax)
								{
									MessageBeep(MB_ICONERROR);
									MessageBox(NULL, L"Min < Max", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
									// Manually changing the text requires invalidating the StringValue property.
									hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
									break;
								}

								g_Variables.UniformMin = newSize;

								hr = PropVariantClear(&var);
							}
							break;
						}
					}
					else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
					{
						UINT selected = ppropvarValue->uintVal;

						switch (selected)
						{
						case 0:
							g_Variables.NormalMean = 0;
							break;
						case UI_COLLECTION_INVALIDINDEX: // The new selection is a value that the user typed.
							if (pCommandExecutionProperties != NULL)
							{
								PROPVARIANT var;
								PropVariantInit(&var);

								hr = pCommandExecutionProperties->GetValue(UI_PKEY_Label, &var); // The text entered by the user is contained in the property set with the pkey UI_PKEY_Label.
								if (FAILED(hr)) { return hr; }

								BSTR bstr = var.bstrVal;
								ULONG newSize = 0UL;

								hr = VarUI4FromStr(bstr,GetUserDefaultLCID(),0,&newSize);
                    
								if (FAILED(hr))
								{
									MessageBeep(MB_ICONERROR);
									MessageBox(NULL, L"VarUI4FromStr Error", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
									// Manually changing the text requires invalidating the StringValue property.
									g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
									break;
								}
								else if (newSize < 0 || newSize > (UINT)pow(2.0, (int)((g_Variables.BitsPerPixel - 8)/3)))
								{
									MessageBeep(MB_ICONERROR);
									MessageBox(NULL, L"0 <= x <= 256", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
									// Manually changing the text requires invalidating the StringValue property.
									g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
									break;
								}

								g_Variables.NormalMean = newSize;

								hr = PropVariantClear(&var);
							}
							break;
						}
					}
				}
			}
			break;
		case cmdHome_EntropyGrid_InitialisationVariable2:
			{
				if (key && *key == UI_PKEY_SelectedItem)
				{
					if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
					{
						UINT selected = ppropvarValue->uintVal;

						switch (selected)
						{
						case 0:
							g_Variables.UniformMax = 0;
							break;
						case UI_COLLECTION_INVALIDINDEX: // The new selection is a value that the user typed.
							if (pCommandExecutionProperties != NULL)
							{
								PROPVARIANT var;
								PropVariantInit(&var);

								hr = pCommandExecutionProperties->GetValue(UI_PKEY_Label, &var); // The text entered by the user is contained in the property set with the pkey UI_PKEY_Label.
								if (FAILED(hr)) { return hr; }

								BSTR bstr = var.bstrVal;
								ULONG newSize = 0UL;

								hr = VarUI4FromStr(bstr, GetUserDefaultLCID(), 0UL, &newSize);
								
								if (FAILED(hr))
								{
									MessageBeep(MB_ICONERROR);
									MessageBox(NULL, L"VarUI4FromStr Error", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
									// Manually changing the text requires invalidating the StringValue property.
									hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
									break;
								}
								else if (newSize < 0 || newSize > (UINT)pow(2.0, (int)((g_Variables.BitsPerPixel - 8)/3)))
								{
									MessageBeep(MB_ICONERROR);
									MessageBox(NULL, L"0 <= x <= 256", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
									// Manually changing the text requires invalidating the StringValue property.
									hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
									break;
								}
								else if (newSize < g_Variables.UniformMin)
								{
									MessageBeep(MB_ICONERROR);
									MessageBox(NULL, L"Max > Min", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
									// Manually changing the text requires invalidating the StringValue property.
									hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
									break;
								}

								g_Variables.UniformMax = newSize;

								hr = PropVariantClear(&var);
							}
							break;
						}
					}
					else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
					{
						UINT selected = ppropvarValue->uintVal;

						switch (selected)
						{
						case 0:
							g_Variables.NormalVariance = 0;
							break;
						case UI_COLLECTION_INVALIDINDEX: // The new selection is a value that the user typed.
							if (pCommandExecutionProperties != NULL)
							{
								PROPVARIANT var;
								PropVariantInit(&var);

								hr = pCommandExecutionProperties->GetValue(UI_PKEY_Label, &var); // The text entered by the user is contained in the property set with the pkey UI_PKEY_Label.
								if (FAILED(hr)) { return hr; }
								
								hr = VarR8FromStr(var.bstrVal, GetUserDefaultLCID(), 0UL, &g_Variables.NormalVariance);
								if (FAILED(hr)) { return hr; }

								hr = PropVariantClear(&var);
							}
							break;
						}
					}
				}
			}
			break;
		case cmdHome_EntropyGrid_Dissipation:
			{
				if (key && *key == UI_PKEY_SelectedItem)
				{
					UINT selected = ppropvarValue->uintVal;

					g_Variables.Dissipation = (DISSIPATION)selected;
				}
			}
			break;
		case cmdHome_LifeGrid_Population:
			{
				if (key && *key == UI_PKEY_SelectedItem)
				{
					UINT selected = ppropvarValue->uintVal;
					switch (selected)
					{
					case 0:
						g_Variables.PopulationInitial = g_Variables.PopulationDefault;
						break;
					case UI_COLLECTION_INVALIDINDEX: // The new selection is a value that the user typed.
						if (pCommandExecutionProperties != NULL)
						{
							PROPVARIANT var;
							PropVariantInit(&var);

							hr = pCommandExecutionProperties->GetValue(UI_PKEY_Label, &var); // The text entered by the user is contained in the property set with the pkey UI_PKEY_Label.
							if (FAILED(hr)) { return hr; }

							BSTR bstr = var.bstrVal;
							ULONG newSize = 0UL;

							hr = VarUI4FromStr(bstr, GetUserDefaultLCID(), 0, &newSize);
                    
							if (FAILED(hr))
							{
								MessageBeep(MB_ICONERROR);
								MessageBox(NULL, L"VarUI4FromStr Error", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
								hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_LifeGrid_Population, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
								break;
							}
							else if (newSize > (ULONG)(g_Variables.Width*g_Variables.Height))
							{
								newSize = (ULONG)(g_Variables.Width*g_Variables.Height);
								hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_LifeGrid_Population, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
							}

							g_Variables.PopulationInitial = newSize;

							hr = PropVariantClear(&var);
						}
						break;
					}
				}
			}
			break;
		case cmdHome_LifeGrid_ReproductionRate:
			{
				if (key && *key == UI_PKEY_SelectedItem)
				{
					UINT selected = ppropvarValue->uintVal;

					switch (selected)
					{
					case 0:
						g_Variables.ReproductionRate = 0.0;
						break;
					case UI_COLLECTION_INVALIDINDEX: // The new selection is a value that the user typed.
						if (pCommandExecutionProperties)
						{
							PROPVARIANT var;
							PropVariantInit(&var);

							hr = pCommandExecutionProperties->GetValue(UI_PKEY_Label, &var); // The text entered by the user is contained in the property set with the pkey UI_PKEY_Label.
							if (FAILED(hr)) { return hr; }

							double newReproductionRate = 0.0;

							hr = VarR8FromStr(var.bstrVal, GetUserDefaultLCID(), 0UL, &newReproductionRate);

							if (FAILED(hr))
							{
								MessageBeep(MB_ICONERROR);
								MessageBox(NULL, L"Reproduction Rate", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
								hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_LifeGrid_ReproductionRate, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
								break;
							}
							else if (newReproductionRate > 1.0)
							{
								newReproductionRate = 1.0;
								hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_LifeGrid_ReproductionRate, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
							}
							else if (newReproductionRate < 0.0)
							{
								newReproductionRate = 0.0;
								hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_LifeGrid_ReproductionRate, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
							}

							g_Variables.ReproductionRate = newReproductionRate;

							hr = PropVariantClear(&var);
						}
						break;
					}
				}
			}
			break;
		case cmdHome_LifeGrid_DeathRate:
			{
				if (key && *key == UI_PKEY_SelectedItem)
				{
					UINT selected = ppropvarValue->uintVal;

					switch (selected)
					{
					case 0:
						g_Variables.DeathRate = 0.0;
						break;
					case UI_COLLECTION_INVALIDINDEX: // The new selection is a value that the user typed.
						if (pCommandExecutionProperties != NULL)
						{
							PROPVARIANT var;
							PropVariantInit(&var);

							hr = pCommandExecutionProperties->GetValue(UI_PKEY_Label, &var); // The text entered by the user is contained in the property set with the pkey UI_PKEY_Label.
							if (FAILED(hr)) { return hr; }

							double newDeathRate = 0.0;
							
							hr = VarR8FromStr(var.bstrVal, GetUserDefaultLCID(), 0UL, &newDeathRate);

							if (FAILED(hr))
							{
								MessageBeep(MB_ICONERROR);
								MessageBox(NULL, L"Death Rate", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
								// Manually changing the text requires invalidating the StringValue property.
								hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_LifeGrid_DeathRate, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
								break;
							}
							else if (newDeathRate > 1.0)
							{
								newDeathRate = 1.0;
								hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_LifeGrid_DeathRate, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
							}
							else if (newDeathRate < 0.0)
							{
								newDeathRate = 0.0;
								hr = g_pRibbonFramework->InvalidateUICommand(cmdHome_LifeGrid_DeathRate, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
							}

							g_Variables.DeathRate = newDeathRate;

							hr = PropVariantClear(&var);
						}
						break;
					}
				}
			}
			break;

		case cmdPaint_Options_Layer:
			{
				if (key && *key == UI_PKEY_SelectedItem)
				{
					UINT selected = ppropvarValue->uintVal;

					g_Variables.PaintLayer = (LAYER)selected;
				}
			}
			break;
        }
    }
    return hr;
}

//
//  FUNCTION: UpdateProperty()
//
//  PURPOSE: Called by the Ribbon framework when a command property (PKEY) needs to be updated.
//
//  COMMENTS:
//
//
//
//
STDMETHODIMP CRibbonHandler::UpdateProperty(
	UINT nCmdID,
    __in REFPROPERTYKEY key,
    __in_opt const PROPVARIANT* ppropvarCurrentValue,
    __out PROPVARIANT* ppropvarNewValue
	)
{

	HRESULT hr = E_FAIL;

	switch(nCmdID)
	{
	case cmdHome_Options_Visualisation:
		{
			if (key == UI_PKEY_BooleanValue)
			{
				hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, g_Variables.Visualisation, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
		}
		break;
	case cmdHome_Options_OutputToFile:
		{
			if (key == UI_PKEY_BooleanValue)
			{
				hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, g_Variables.OutputToFile, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
		}
		break;
	case cmdHome_Options_Paint:
		{
			if (key == UI_PKEY_BooleanValue)
			{
				hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, g_Variables.Paint, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
		}
		break;
	case cmdHome_Commands_Run:
		{
			if (key == UI_PKEY_Enabled)
			{
				hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, (bool)(g_Variables.EntropyInitialised && g_Variables.Population), ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
			else if (key == UI_PKEY_BooleanValue)
			{
				hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, g_Variables.Running, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
			else if (key == UI_PKEY_Label) //Update the Label of Run control
			{
				if (g_Variables.Running)
				{
					hr = UIInitPropertyFromString(key, L"Stop", ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
				else
				{
					hr = UIInitPropertyFromString(key, L"Run", ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
			}
			else if (key == UI_PKEY_LargeImage)
			{					
				if (g_Variables.Running)
				{
					if (!_IUIImageStopLarge)
					{
						hr = CreateUIImageFromBitmapResource(MAKEINTRESOURCEW(cmdHome_Commands_Stop_LargeImages_96__RESID), &_IUIImageStopLarge);
						if (FAILED(hr)) { return hr; }
					}

					hr = UIInitPropertyFromImage(key, _IUIImageStopLarge, ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
				else
				{
					if (!_IUIImageRunLarge)
					{
						hr = CreateUIImageFromBitmapResource(MAKEINTRESOURCEW(cmdHome_Commands_Run_LargeImages_96__RESID), &_IUIImageRunLarge);
						if (FAILED(hr)) { return hr; }
					}

					hr = UIInitPropertyFromImage(key, _IUIImageRunLarge, ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
			}
			else if (key == UI_PKEY_SmallImage)
			{					
				if (g_Variables.Running)
				{
					if (!_IUIImageStopSmall)
					{
						hr = CreateUIImageFromBitmapResource(MAKEINTRESOURCEW(cmdHome_Commands_Stop_SmallImages_96__RESID), &_IUIImageStopSmall);
						if (FAILED(hr)) { return hr; }
					}

					hr = UIInitPropertyFromImage(key, _IUIImageStopSmall, ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
				else
				{
					if (!_IUIImageRunSmall)
					{
						hr = CreateUIImageFromBitmapResource(MAKEINTRESOURCEW(cmdHome_Commands_Run_SmallImages_96__RESID), &_IUIImageRunSmall);
						if (FAILED(hr)) { return hr; }
					}

					hr = UIInitPropertyFromImage(key, _IUIImageRunSmall, ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
			}
		}
		break;
	case cmdHome_Commands_NextStep:
		{
			if (key == UI_PKEY_Enabled)
			{
				hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, (bool)(g_Variables.EntropyInitialised && g_Variables.Population), ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
		}
		break;
	case cmdHome_Commands_Clear:
		{
			if (key == UI_PKEY_Enabled)
			{
				hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, (bool)(g_Variables.EntropyInitialised || g_Variables.Population), ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
		}
		break;
	case cmdHome_Commands_ClearEntropyGrid:
		{
			if (key == UI_PKEY_Enabled)
			{
				hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, g_Variables.EntropyInitialised, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
		}
		break;
	case cmdHome_Commands_ClearLifeGrid:
		{
			if (key == UI_PKEY_Enabled)
			{
				hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, (g_Variables.Population > 0)? true : false, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
		}
		break;
	case cmdHome_Ruleset_RandomNumberFunction:
		{
			if (key == UI_PKEY_ItemsSource)
			{
				Microsoft::WRL::ComPtr<IUICollection> pCollection;

				hr = ppropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				if (FAILED(hr)) { return hr; }

				int labelIds[] = {RANDOM_NUMBER_FUNCTION_1, RANDOM_NUMBER_FUNCTION_2};

				// Populate the combobox
				for (int i = 0; i < _countof(labelIds) && SUCCEEDED(hr); i++)
				{
					// Create a new property set for each item.
					Microsoft::WRL::ComPtr<CRibbonSimplePropertySet> pItem;

					hr = CRibbonSimplePropertySet::CreateInstance(&pItem);
					if (FAILED(hr)) { return hr; }
  
					// Load the label from the resource file.
					WCHAR wszLabel[MAX_RESOURCE_LENGTH];
					(void)LoadStringW(GetModuleHandleW(nullptr), labelIds[i], wszLabel, MAX_RESOURCE_LENGTH);

					// Initialize the property set with no image, the label that was just loaded, and no category.
					pItem->InitializeItemProperties(nullptr, wszLabel, UI_COLLECTION_INVALIDINDEX);

					hr = pCollection->Add(pItem.Get());
				}
			}
			else if (key == UI_PKEY_Categories)
			{
				// A return value of S_FALSE or E_NOTIMPL will result in a gallery with no categories.
				// If you return any error other than E_NOTIMPL, the contents of the gallery will not display.
				hr = S_FALSE;
			}
			else if (key == UI_PKEY_SelectedItem)
			{
				// The default selection of item.
				hr = UIInitPropertyFromUInt32(UI_PKEY_SelectedItem, g_Variables.RandomNumberFunction, ppropvarNewValue);
			}
		}
		break;
	case cmdHome_Ruleset_PeriodicBoundaryConditions:
		{
			if (key == UI_PKEY_BooleanValue)
			{
				hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, g_Variables.PeriodicBoundaryConditions, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
			else if (key == UI_PKEY_Enabled)
			{
				hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, false, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
		}
		break;
	case cmdHome_EntropyGrid_Initialisation:
		{
			if (key == UI_PKEY_ItemsSource)
			{
				Microsoft::WRL::ComPtr<IUICollection> pCollection;

				hr = ppropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				if (FAILED(hr)) { return hr; }

				int labelIds[] = {INITIALISATION_1, INITIALISATION_2};

				// Populate the combobox with the three layout options.
				for (int i = 0; i < _countof(labelIds) && SUCCEEDED(hr); i++)
				{
					// Create a new property set for each item.
					Microsoft::WRL::ComPtr<CRibbonSimplePropertySet> pItem;

					hr = CRibbonSimplePropertySet::CreateInstance(&pItem);
					if (FAILED(hr)) { return hr; }
					
					// Load the label from the resource file.
					WCHAR wszLabel[MAX_RESOURCE_LENGTH];
					(void)LoadStringW(GetModuleHandleW(nullptr), labelIds[i], wszLabel, MAX_RESOURCE_LENGTH);

					// Initialize the property set with no image, the label that was just loaded, and no category.
					pItem->InitializeItemProperties(nullptr, wszLabel, UI_COLLECTION_INVALIDINDEX);

					hr = pCollection->Add(pItem.Get());
				}
			}
			else if (key == UI_PKEY_Categories)
			{
				// A return value of S_FALSE or E_NOTIMPL will result in a gallery with no categories.
				// If you return any error other than E_NOTIMPL, the contents of the gallery will not display.
				hr = S_FALSE;
			}
			else if (key == UI_PKEY_SelectedItem)
			{
				// The default selection of item.
				hr = UIInitPropertyFromUInt32(UI_PKEY_SelectedItem, g_Variables.Initialisation, ppropvarNewValue);
			}
		}
		break;
	case cmdHome_EntropyGrid_InitialisationVariable1:
		{
			if (key == UI_PKEY_ItemsSource)
			{
				Microsoft::WRL::ComPtr<IUICollection> pCollection;

				hr = ppropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				if (FAILED(hr)) { return hr; }

				// Create a new property set for each item.
				Microsoft::WRL::ComPtr<CRibbonSimplePropertySet> pItem;

				hr = CRibbonSimplePropertySet::CreateInstance(&pItem);
				if (FAILED(hr)) { return hr; }

				// Initialize the property set with no image, the label that was just loaded, and no category.
				pItem->InitializeItemProperties(nullptr, L"0", UI_COLLECTION_INVALIDINDEX);

				hr = pCollection->Add(pItem.Get());
			}
			else if (key == UI_PKEY_Label)
			{
				if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
				{
					hr = UIInitPropertyFromString(key, L"Min", ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
				else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
				{
					hr = UIInitPropertyFromString(key, L"Mean", ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
			}
			else if (key == UI_PKEY_Categories)
			{
				// A return value of S_FALSE or E_NOTIMPL will result in a gallery with no categories.
				// If you return any error other than E_NOTIMPL, the contents of the gallery will not display.
				hr = S_FALSE;
			}
			else if (key == UI_PKEY_SelectedItem)
			{
				// The default selection of item.
				hr = UIInitPropertyFromUInt32(UI_PKEY_SelectedItem, 0, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
			else if (key == UI_PKEY_StringValue)
			{
				BSTR bstr = nullptr;

				if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
				{
					hr = VarBstrFromUI4(g_Variables.UniformMin, GetUserDefaultLCID(), 0, &bstr);
					if (FAILED(hr)) { return hr; }
				}
				else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
				{
					hr = VarBstrFromUI4(g_Variables.NormalMean, GetUserDefaultLCID(), 0, &bstr);
					if (FAILED(hr)) { return hr; }
				}
        
				hr = UIInitPropertyFromString(UI_PKEY_StringValue, bstr, ppropvarNewValue);

				SysFreeString(bstr);

				if (FAILED(hr)) { return hr; }
			}
			else if (key == UI_PKEY_TooltipTitle)
			{
				if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
				{
					hr = UIInitPropertyFromString(key, L"Min", ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
				else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
				{
					hr = UIInitPropertyFromString(key, L"Mean", ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
			}
			else if (key == UI_PKEY_TooltipDescription)
			{
				if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
				{
					hr = UIInitPropertyFromString(key, L"Min", ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
				else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
				{
					hr = UIInitPropertyFromString(key, L"Mean", ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
			}
			else if (key == UI_PKEY_LargeImage)
			{					
				if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
				{
					if (!_IUIImageMinLarge)
					{
						hr = CreateUIImageFromBitmapResource(MAKEINTRESOURCE(cmdHome_EntropyGrid_InitialisationVariable1_LargeImages_96__RESID), &_IUIImageMinLarge);
						if (FAILED(hr)) { return hr; }
					}

					hr = UIInitPropertyFromImage(key, _IUIImageMinLarge, ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
				else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
				{
					if (!_IUIImageMeanLarge)
					{
						hr = CreateUIImageFromBitmapResource(MAKEINTRESOURCE(cmdHome_EntropyGrid_Mean_LargeImages_96__RESID), &_IUIImageMeanLarge);
						if (FAILED(hr)) { return hr; }
					}

					hr = UIInitPropertyFromImage(key, _IUIImageMeanLarge, ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
			}
			else if (key == UI_PKEY_SmallImage)
			{					
				if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
				{
					if (!_IUIImageMinSmall)
					{
						hr = CreateUIImageFromBitmapResource(MAKEINTRESOURCE(cmdHome_EntropyGrid_InitialisationVariable1_SmallImages_96__RESID), &_IUIImageMinSmall);
						if (FAILED(hr)) { return hr; }
					}

					hr = UIInitPropertyFromImage(key, _IUIImageMinSmall, ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
				else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
				{
					if (!_IUIImageMeanSmall)
					{
						hr = CreateUIImageFromBitmapResource(MAKEINTRESOURCE(cmdHome_EntropyGrid_Mean_SmallImages_96__RESID), &_IUIImageMeanSmall);
						if (FAILED(hr)) { return hr; }
					}

					hr = UIInitPropertyFromImage(key, _IUIImageMeanSmall, ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
			}
		}
		break;
	case cmdHome_EntropyGrid_InitialisationVariable2:
		{
			if (key == UI_PKEY_ItemsSource)
			{
				Microsoft::WRL::ComPtr<IUICollection> pCollection;

				hr = ppropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				if (FAILED(hr)) { return hr; }

				// Create a new property set for each item.
				Microsoft::WRL::ComPtr<CRibbonSimplePropertySet> pItem;

				hr = CRibbonSimplePropertySet::CreateInstance(&pItem);
				if (FAILED(hr)) { return hr; }

				// Initialize the property set with no image, the label that was just loaded, and no category.
				pItem->InitializeItemProperties(nullptr, L"0", UI_COLLECTION_INVALIDINDEX);

				hr = pCollection->Add(pItem.Get());
			}
			else if (key == UI_PKEY_Label)
			{
				if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
				{
					hr = UIInitPropertyFromString(key, L"Max", ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
				else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
				{
					hr = UIInitPropertyFromString(key, L"Variance", ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
			}
			else if (key == UI_PKEY_Categories)
			{
				// A return value of S_FALSE or E_NOTIMPL will result in a gallery with no categories.
				// If you return any error other than E_NOTIMPL, the contents of the gallery will not display.
				hr = S_FALSE;
			}
			else if (key == UI_PKEY_SelectedItem)
			{
				// The default selection of item.
				hr = UIInitPropertyFromUInt32(UI_PKEY_SelectedItem, 0, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
			else if (key == UI_PKEY_StringValue)
			{
				if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
				{
					BSTR bstr;

					hr = VarBstrFromUI4(g_Variables.UniformMax, GetUserDefaultLCID(), 0, &bstr);
					if (FAILED(hr)) { return hr; }

					hr = UIInitPropertyFromString(UI_PKEY_StringValue, bstr, ppropvarNewValue);
					SysFreeString(bstr);
					if (FAILED(hr)) { return hr; }
				}
				else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
				{
					wchar_t wNormalVariance[MAX_RESOURCE_LENGTH];

					hr = StringCchPrintfW(wNormalVariance, MAX_RESOURCE_LENGTH, L"%.6f", g_Variables.NormalVariance);
					if (FAILED(hr)) { return hr; }

					hr = UIInitPropertyFromString(UI_PKEY_StringValue, wNormalVariance, ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
			}
			else if (key == UI_PKEY_TooltipTitle)
			{
				if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
				{
					hr = UIInitPropertyFromString(key, L"Max", ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
				else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
				{
					hr = UIInitPropertyFromString(key, L"Variance", ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
			}
			else if (key == UI_PKEY_TooltipDescription)
			{
				if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
				{
					hr = UIInitPropertyFromString(key, L"Max", ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
				else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
				{
					hr = UIInitPropertyFromString(key, L"Variance", ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
			}
			else if (key == UI_PKEY_LargeImage)
			{					
				if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
				{
					if (!_IUIImageMaxLarge)
					{
						hr = CreateUIImageFromBitmapResource(MAKEINTRESOURCE(cmdHome_EntropyGrid_InitialisationVariable2_LargeImages_96__RESID), &_IUIImageMaxLarge);
						if (FAILED(hr)) { return hr; }
					}

					hr = UIInitPropertyFromImage(key, _IUIImageMaxLarge, ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
				else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
				{
					if (!_IUIImageVarianceLarge)
					{
						hr = CreateUIImageFromBitmapResource(MAKEINTRESOURCE(cmdHome_EntropyGrid_Variance_LargeImages_96__RESID), &_IUIImageVarianceLarge);
						if (FAILED(hr)) { return hr; }
					}

					hr = UIInitPropertyFromImage(key, _IUIImageVarianceLarge, ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
			}
			else if (key == UI_PKEY_SmallImage)
			{					
				if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
				{
					if (!_IUIImageMaxSmall)
					{
						hr = CreateUIImageFromBitmapResource(MAKEINTRESOURCE(cmdHome_EntropyGrid_InitialisationVariable2_SmallImages_96__RESID), &_IUIImageMaxSmall);
						if (FAILED(hr)) { return hr; }
					}

					hr = UIInitPropertyFromImage(key, _IUIImageMaxSmall, ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
				else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
				{
					if (!_IUIImageVarianceSmall)
					{
						hr = CreateUIImageFromBitmapResource(MAKEINTRESOURCE(cmdHome_EntropyGrid_Variance_SmallImages_96__RESID), &_IUIImageVarianceSmall);
						if (FAILED(hr)) { return hr; }
					}

					hr = UIInitPropertyFromImage(key, _IUIImageVarianceSmall, ppropvarNewValue);
					if (FAILED(hr)) { return hr; }
				}
			}
		}
		break;
	case cmdHome_EntropyGrid_Dissipation:
		{
			if (key == UI_PKEY_ItemsSource)
			{
				Microsoft::WRL::ComPtr<IUICollection> pCollection;

				hr = ppropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				if (FAILED(hr)) { return hr; }

				int labelIds[] = {DISSIPATION_1, DISSIPATION_2, DISSIPATION_3};

				// Populate the combobox with the three layout options.
				for (int i=0; i<_countof(labelIds) && SUCCEEDED(hr); i++)
				{
					// Create a new property set for each item.
					Microsoft::WRL::ComPtr<CRibbonSimplePropertySet> pItem;

					hr = CRibbonSimplePropertySet::CreateInstance(&pItem);
					if (FAILED(hr)) { return hr; }
					
					// Load the label from the resource file.
					WCHAR wszLabel[MAX_RESOURCE_LENGTH];
					(void)LoadStringW(GetModuleHandleW(nullptr), labelIds[i], wszLabel, MAX_RESOURCE_LENGTH);

					// Initialize the property set with no image, the label that was just loaded, and no category.
					pItem->InitializeItemProperties(nullptr, wszLabel, UI_COLLECTION_INVALIDINDEX);

					hr = pCollection->Add(pItem.Get());
				}
			}
			else if (key == UI_PKEY_Categories)
			{
				// A return value of S_FALSE or E_NOTIMPL will result in a gallery with no categories.
				// If you return any error other than E_NOTIMPL, the contents of the gallery will not display.
				hr = S_FALSE;
			}
			else if (key == UI_PKEY_SelectedItem)
			{
				// The default selection of item.
				hr = UIInitPropertyFromUInt32(UI_PKEY_SelectedItem, g_Variables.Dissipation, ppropvarNewValue);
			}
		}
		break;
	case cmdHome_LifeGrid_Population:
		{
			if (key == UI_PKEY_ItemsSource)
			{
				Microsoft::WRL::ComPtr<IUICollection> pCollection;

				hr = ppropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				if (FAILED(hr)) { return hr; }

				// Create a new property set for each item.
				Microsoft::WRL::ComPtr<CRibbonSimplePropertySet> pItem;

				hr = CRibbonSimplePropertySet::CreateInstance(&pItem);
				if (FAILED(hr)) { return hr; }

				wchar_t wPopulationDefault[MAX_RESOURCE_LENGTH];

				hr = StringCchPrintfW(wPopulationDefault, MAX_RESOURCE_LENGTH, L"%d", g_Variables.PopulationDefault);
				if (FAILED(hr)) { return hr; }

				// Initialize the property set with no image, the label that was just loaded, and no category.
				pItem->InitializeItemProperties(nullptr, wPopulationDefault, UI_COLLECTION_INVALIDINDEX);

				hr = pCollection->Add(pItem.Get());
			}
			else if (key == UI_PKEY_Categories)
			{
				// A return value of S_FALSE or E_NOTIMPL will result in a gallery with no categories.
				// If you return any error other than E_NOTIMPL, the contents of the gallery will not display.
				hr = S_FALSE;
			}
			else if (key == UI_PKEY_SelectedItem)
			{
				// The default selection of item.
				hr = UIInitPropertyFromUInt32(UI_PKEY_SelectedItem, 0, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
			else if (key == UI_PKEY_StringValue)
			{
				BSTR bstr = nullptr;

				hr = VarBstrFromUI4(g_Variables.PopulationInitial, GetUserDefaultLCID(), 0, &bstr);
				if (FAILED(hr)) { return hr; }
        
				hr = UIInitPropertyFromString(UI_PKEY_StringValue, bstr, ppropvarNewValue);

				SysFreeString(bstr);

				if (FAILED(hr)) { return hr; }
			}
		}
		break;
	case cmdHome_LifeGrid_ReproductionRate:
		{
			if (key == UI_PKEY_ItemsSource)
			{
				Microsoft::WRL::ComPtr<IUICollection> pCollection;

				hr = ppropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				if (FAILED(hr)) { return hr; }

				// Create a new property set for each item.
				Microsoft::WRL::ComPtr<CRibbonSimplePropertySet> pItem;

				hr = CRibbonSimplePropertySet::CreateInstance(&pItem);
				if (FAILED(hr)) { return hr; }

				wchar_t wReproductionRate[MAX_RESOURCE_LENGTH];

				hr = StringCchPrintfW(wReproductionRate, MAX_RESOURCE_LENGTH, L"%.6f", g_Variables.ReproductionRate);
				if (FAILED(hr)) { return hr; }

				// Initialize the property set with no image, the label that was just loaded, and no category.
				pItem->InitializeItemProperties(nullptr, wReproductionRate, UI_COLLECTION_INVALIDINDEX);

				hr = pCollection->Add(pItem.Get());
			}
			else if (key == UI_PKEY_Categories)
			{
				// A return value of S_FALSE or E_NOTIMPL will result in a gallery with no categories.
				// If you return any error other than E_NOTIMPL, the contents of the gallery will not display.
				hr = S_FALSE;
			}
			else if (key == UI_PKEY_SelectedItem)
			{
				// The default selection of item.
				hr = UIInitPropertyFromUInt32(UI_PKEY_SelectedItem, 0, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
			else if (key == UI_PKEY_StringValue)
			{
				wchar_t wReproductionRate[MAX_RESOURCE_LENGTH];

				hr = StringCchPrintfW(wReproductionRate, MAX_RESOURCE_LENGTH, L"%.6f", g_Variables.ReproductionRate);
				if (FAILED(hr)) { return hr; }
        
				hr = UIInitPropertyFromString(UI_PKEY_StringValue, wReproductionRate, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
		}
		break;
	case cmdHome_LifeGrid_DeathRate:
		{
			if (key == UI_PKEY_ItemsSource)
			{
				Microsoft::WRL::ComPtr<IUICollection> pCollection;

				hr = ppropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				if (FAILED(hr)) { return hr; }

				// Create a new property set for each item.
				Microsoft::WRL::ComPtr<CRibbonSimplePropertySet> pItem;

				hr = CRibbonSimplePropertySet::CreateInstance(&pItem);
				if (FAILED(hr)) { return hr; }

				wchar_t wDeathRate[MAX_RESOURCE_LENGTH];

				hr = StringCchPrintfW(wDeathRate, MAX_RESOURCE_LENGTH, L"%.6f", g_Variables.DeathRate);
				if (FAILED(hr)) { return hr; }

				// Initialize the property set with no image, the label that was just loaded, and no category.
				pItem->InitializeItemProperties(nullptr, wDeathRate, UI_COLLECTION_INVALIDINDEX);

				hr = pCollection->Add(pItem.Get());
			}
			else if (key == UI_PKEY_Categories)
			{
				// A return value of S_FALSE or E_NOTIMPL will result in a gallery with no categories.
				// If you return any error other than E_NOTIMPL, the contents of the gallery will not display.
				hr = S_FALSE;
			}
			else if (key == UI_PKEY_SelectedItem)
			{
				// The default selection of item.
				hr = UIInitPropertyFromUInt32(UI_PKEY_SelectedItem, 0, ppropvarNewValue);
				if (FAILED(hr))
				{
					return hr;
				}
			}
			else if (key == UI_PKEY_StringValue)
			{
				wchar_t wDeathRate[MAX_RESOURCE_LENGTH];

				hr = StringCchPrintfW(wDeathRate, MAX_RESOURCE_LENGTH, L"%.6f", g_Variables.DeathRate);
				if (FAILED(hr)) { return hr; }
        
				hr = UIInitPropertyFromString(UI_PKEY_StringValue, wDeathRate, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
		}
		break;
	case cmdHome_LifeGrid_Intelligence:
		{
			if (key == UI_PKEY_BooleanValue)
			{
				hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, g_Variables.Intelligence, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
		}
		break;
	case cmdHome_LifeGrid_Bias:
		{
			if (key == UI_PKEY_BooleanValue)
			{
				hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, g_Variables.Bias, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
		}
		break;
	case cmdContextualTabGroup_Paint:
		{
			if (key == UI_PKEY_ContextAvailable)
			{
				UI_CONTEXTAVAILABILITY contextAvailabilityPaint;

				if (g_Variables.Paint)
				{
					contextAvailabilityPaint = UI_CONTEXTAVAILABILITY_AVAILABLE;
				}
				else
				{
					contextAvailabilityPaint = UI_CONTEXTAVAILABILITY_NOTAVAILABLE;
				}

				hr = UIInitPropertyFromUInt32(key, contextAvailabilityPaint, ppropvarNewValue);
				if (FAILED(hr)) { return hr; }
			}
		}
		break;
	case cmdPaint_Options_Layer:
		{
			if (key == UI_PKEY_ItemsSource)
			{
				Microsoft::WRL::ComPtr<IUICollection> pCollection;

				hr = ppropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				if (FAILED(hr)) { return hr; }

				int labelIds[] = {LAYER_1, LAYER_2};

				// Populate the combobox with the three layout options.
				for (int i = 0; i < _countof(labelIds) && SUCCEEDED(hr); i++)
				{
					// Create a new property set for each item.
					Microsoft::WRL::ComPtr<CRibbonSimplePropertySet> pItem;

					hr = CRibbonSimplePropertySet::CreateInstance(&pItem);
					if (FAILED(hr)) { return hr; }
  
					// Load the label from the resource file.
					WCHAR wszLabel[MAX_RESOURCE_LENGTH];
					(void)LoadStringW(GetModuleHandleW(nullptr), labelIds[i], wszLabel, MAX_RESOURCE_LENGTH);

					// Initialize the property set with no image, the label that was just loaded, and no category.
					pItem->InitializeItemProperties(nullptr, wszLabel, UI_COLLECTION_INVALIDINDEX);

					hr = pCollection->Add(pItem.Get());
				}
			}
			else if (key == UI_PKEY_Categories)
			{
				// A return value of S_FALSE or E_NOTIMPL will result in a gallery with no categories.
				// If you return any error other than E_NOTIMPL, the contents of the gallery will not display.
				hr = S_FALSE;
			}
			else if (key == UI_PKEY_SelectedItem)
			{
				// The default selection of item.
				hr = UIInitPropertyFromUInt32(UI_PKEY_SelectedItem, g_Variables.PaintLayer, ppropvarNewValue);
			}
		}
		break;
	}

	return hr;
}

HRESULT CRibbonHandler::CreateInstance(__deref_out CRibbonHandler **ppHandler)
{
    if (!ppHandler)
    {
        return E_POINTER;
    }

    *ppHandler = nullptr;

    HRESULT hr = S_OK;

    CRibbonHandler* pHandler = new CRibbonHandler();

    if (pHandler)
    {
        *ppHandler = pHandler;
    }
    else
    {
        hr = E_OUTOFMEMORY;
    }

    return hr;
}

// IUnknown methods.
STDMETHODIMP_(ULONG) CRibbonHandler::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CRibbonHandler::Release()
{
    LONG cRef = InterlockedDecrement(&m_cRef);
    if (cRef == 0L)
    {
        delete this;
    }

    return cRef;
}

STDMETHODIMP CRibbonHandler::QueryInterface(REFIID iid, void** ppv)
{
    if (!ppv)
    {
        return E_POINTER;
    }

    if (iid == __uuidof(IUnknown))
    {
        *ppv = static_cast<IUnknown*>(this);
    }
    else if (iid == __uuidof(IUICommandHandler))
    {
        *ppv = static_cast<IUICommandHandler*>(this);
    }
    else 
    {
        *ppv = NULL;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}