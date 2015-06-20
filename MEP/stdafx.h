// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#define _CRT_RAND_S // rand_s

// Section below required for TaskDialogIndirect
#if defined _M_IX86
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_IA64
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='ia64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif

#include "targetver.h"

#define STRICT
//#define _CRT_SECURE_CPP_OVERLOAD_STANDARD_NAMES 1
//#define _CRT_SECURE_CPP_OVERLOAD_STANDARD_NAMES_COUNT 1
// Windows Header Files:
#include <windows.h>
#include <Windowsx.h> // GET_X_LPARAM, GET_Y_LPARAM

// C RunTime Header Files
#include <stdlib.h>
#include <tchar.h> // _tcsrchr, _tcsicmp

// Additional headers
#include <commctrl.h> // TASKDIALOGCONFIG
#include <list> // std::list
#include <strsafe.h> // StringCchCopy, StringCchPrintf
#include <shellapi.h> // HDROP, DragQueryFile
#include <ppl.h> // Concurrency::Parallel_for
#include <wrl\client.h> // Microsoft::WRL::ComPtr

enum INITIALISATION
{
	INITIALISATION_NORMAL = 0,
    INITIALISATION_UNIFORM = 1
};

enum DISSIPATION
{
    DISSIPATION_NONE = 0,
    DISSIPATION_AVERAGECEIL = 1,
    DISSIPATION_AVERAGEFLOOR = 2
};

enum LAYER
{
    LAYER_ENTROPY = 0,
    LAYER_LIFE = 1
};

enum RANDOM_NUMBER_FUNCTION
{
    RAN = 0,
    RAND_S = 1
};

enum SEX
{
	MALE = 0,
	FEMALE = 1
};

struct LIFEPROPERTIES
{
	UINT x;
	UINT y;
	SEX sex;
	UINT Lifespan;
};

struct VARIABLES
{
	bool					NewRun;
	UINT					Width;
	UINT					Height;
	UINT					BitsPerPixel;
	UINT					BytesPerPixel;
	UINT					Stride;
    bool					Visualisation;
    bool					OutputToFile;
    bool					Paint;
	bool					Redraw;
	bool					Running;
	UINT					NumberOfSteps;
	RANDOM_NUMBER_FUNCTION	RandomNumberFunction;
    bool					PeriodicBoundaryConditions;
	INITIALISATION			Initialisation;
	UINT					UniformMin;
	UINT					UniformMax;
	UINT					NormalMean;
	double					NormalVariance;
	DISSIPATION				Dissipation;
	long					PopulationDefault;
	long					PopulationInitial;
	long					Population;
	double					ReproductionRate;
	double					DeathRate;
    bool					Intelligence;
    bool					Bias;
	LAYER					PaintLayer;
	bool					EntropyInitialised;
	UINT					**EntropyCurrent;
	UINT					**EntropyNext;
	bool					LifeInitialised;
	byte					*BitmapDataEntropy;
	byte					*BitmapDataLife;
	UINT					Generation;
	double					InformationEntropy;
	UINT					FreeEnergyInitial;
	UINT					FreeEnergy;
	double					ThermodynamicEntropy;
	double					FPS;
	std::list<LIFEPROPERTIES> EcosystemCurrent;
	POINT					TranslateVector;
};

template<class Interface>
inline void SafeRelease(Interface **ppInterfaceToRelease)
{
	if (*ppInterfaceToRelease != NULL)
	{
		(*ppInterfaceToRelease)->Release();

		(*ppInterfaceToRelease) = NULL;
	}
}