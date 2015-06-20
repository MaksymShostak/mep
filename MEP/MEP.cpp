// MEP.cpp : Defines the entry point for the application.

#include "stdafx.h"
#include "MEP.h"
#include "RibbonFramework.h"
#include "Direct2DRenderer.h"
#include "DropTarget.h"
#include "CDPI.h"
#include "HRESULT.h"

#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;								// current instance
WCHAR szTitle[MAX_LOADSTRING];					// The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];			// the main window class name
VARIABLES g_Variables = {0};
Direct2DRenderer renderer;
HWND hWnd = nullptr;
HWND hRenderTarget = nullptr;
HWND hStatusBar = nullptr;
ITaskbarList3 *g_pTaskbarList = nullptr;
WCHAR OutputFileName[MAX_PATH];
std::ofstream myfile;

CDPI g_metrics;

struct Ran
{
	//Implementation of the highest quality recommended generator. The constructor is called with
	//an integer seed and creates an instance of the generator. The member functions int64, doub,
	//and int32 return the next values in the random sequence, as a variable type indicated by their
	//names. The period of the generator is 3.138x10^57.
	unsigned long long u, v, w;

	Ran(unsigned long long j) : v(4101842887655102017LL), w(1)
	{
		//Constructor. Call with any integer seed (except value of v above).
		u = j ^ v; int64();
		v = u; int64();
		w = v; int64();
	}

	inline unsigned long long int64()
	{
		//Return 64-bit random integer. See text for explanation of method.
		u = u * 2862933555777941757LL + 7046029254386353087LL;
		v ^= v >> 17; v ^= v << 31; v ^= v >> 8;
		w = 4294957665U*(w & 0xffffffff) + (w >> 32);
		ULONGLONG x = u ^ (u << 21); x ^= x >> 35; x ^= x << 4;
		return (x + v) ^ w;
	}

	inline double doub() //Return random double-precision floating value in the range 0. to 1.
	{
		return 5.42101086242752217E-20 * int64();
	}
	
	inline UINT int32() //Return 32-bit random integer.
	{
		return (UINT)int64();
	}
};

//Initialise random number generators
Ran ran(time(0));
UINT random_number = 0U;

namespace RandomNumber
{
	double Double(RANDOM_NUMBER_FUNCTION rnf)
	{
		if (rnf == RAN)
		{
			return ran.doub();
		}
		else if (rnf == RAND_S)
		{
			rand_s(&random_number);
			return (double)random_number/((double)UINT_MAX + 1.0);
		}
		return 0;
	}

	UINT Uniform(RANDOM_NUMBER_FUNCTION rnf, UINT Min, UINT Max) // Min <= UINT < Max
	{
		return (UINT)(RandomNumber::Double(rnf) * (Max - Min)) + Min;
	}

	double Normal(RANDOM_NUMBER_FUNCTION rnf, UINT Mean, double Variance)
	{
		double u, v, x, y, q;

		do
		{
			u = RandomNumber::Double(rnf);
			v = 1.7156*(RandomNumber::Double(rnf)-0.5);
			x = u - 0.449871;
			y = abs(v) + 0.386595;
			q = (x*x) + y*(0.19600*y-0.25472*x);
		}
		while (q > 0.27597 && (q > 0.27846 || (v*v) > -4.*log(u)*(u*u)));

		return Mean + Variance*v/u;
	}
}

inline bool OutputToFile()
{
	if (g_Variables.NewRun)
	{
		SYSTEMTIME LocalTime;

   		GetLocalTime(&LocalTime);

		HRESULT hr = StringCchPrintf(
			OutputFileName,
			MAX_PATH,
			TEXT("output (%d-%02d-%02d %02d %02d %02d %03d).csv"),
			LocalTime.wYear, LocalTime.wMonth, LocalTime.wDay, LocalTime.wHour, LocalTime.wMinute, LocalTime.wSecond, LocalTime.wMilliseconds);
		
		if (SUCCEEDED(hr))
		{
			/*if (GetFileAttributesA(OutputFileName) != INVALID_FILE_ATTRIBUTES)
			{
				MessageBeep(MB_ICONWARNING);
				MessageBox(hWnd, L"File already exists.\nPlease wait 1 millisecond.", NULL, MB_ICONWARNING | MB_SETFOREGROUND | MB_APPLMODAL);
			}*/

			myfile.open(OutputFileName, std::ios::out | std::ios::app );

			if (myfile.is_open())
			{
				myfile << "Generation, Population, IEntropy, TEntropy" << std::endl;
						
				g_Variables.NewRun = false;
			}
			else
			{
				MessageBeep(MB_ICONWARNING);
				MessageBox(hWnd, TEXT("Cannot open file for writing.\nOutput to file disabled."), NULL, MB_ICONWARNING | MB_SETFOREGROUND | MB_TASKMODAL);
				g_Variables.OutputToFile = false;
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Options_OutputToFile, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_BooleanValue);

				return false;
			}
		}
	}

	myfile << g_Variables.Generation << ", " << g_Variables.Population << ", ";
	myfile.precision(10);
	myfile << g_Variables.InformationEntropy << ", ";
	myfile << g_Variables.ThermodynamicEntropy;
	myfile << std::endl;

	return true;
}

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE /*hPrevInstance*/, LPTSTR /*lpCmdLine*/, int nCmdShow)
{
	HRESULT hr = S_OK;
	MSG msg = {0};
	HACCEL hAccelTable = nullptr;
	int	nArgs = 0;
	LPWSTR *lpszArgv = nullptr;

	HeapSetInformation(NULL, HeapEnableTerminationOnCorruption, NULL, 0);

	lpszArgv = CommandLineToArgvW(GetCommandLineW(), &nArgs);

	if (lpszArgv == nullptr)
	{
		ErrorDescription(E_POINTER);
		return EXIT_FAILURE;
	}

	for (int i = 1; i < nArgs; i++)
	{
		if (_wcsicmp(lpszArgv[i], L"-width") == 0)
		{
			if (i + 1 < nArgs)
			{
				i++;
				g_Variables.Width = _wtoi(lpszArgv[i]);
			}
		}
		else if (wcscmp(lpszArgv[i], L"-height") == 0)
		{
			if (i + 1 < nArgs)
			{
				i++;
				g_Variables.Height = _wtoi(lpszArgv[i]);
			}
		}
		else
		{
			//wprintf_s(L"Error with command line argument %d: '%ls'\n", i, lpszArgv[i]);
		}
	}

	if (g_Variables.Width == 0)
	{
		if (g_Variables.Height == 0)
		{
			g_Variables.Width = 512;
		}
		else
		{
			g_Variables.Width = g_Variables.Height;
		}
	}

	if (g_Variables.Height == 0)
	{
		if (g_Variables.Width == 0)
		{
			g_Variables.Height = 512;
		}
		else
		{
			g_Variables.Height = g_Variables.Width;
		}
	}

	/*if (nArgs == 2)
	{
		g_Variables.Width = _wtoi(lpszArgv[1]);
		g_Variables.Height = g_Variables.Width;
	}
	else if (nArgs == 3)
	{
		g_Variables.Width = _wtoi(lpszArgv[1]);
		g_Variables.Height = _wtoi(lpszArgv[2]);
	}
	else
	{
		g_Variables.Width = 512;
		g_Variables.Height = 512;
	}

	/*for(int i = 1; i < nArgs; i++) //Start at 1 because 0 is the path
	{
		MessageBoxW(NULL, lpszArgv[i], NULL, MB_OK);
	}*/

	if (LocalFree(lpszArgv)) // If the function succeeds, the return value is NULL. If the function fails, the return value is equal to a handle to the local memory object.
	{
		ErrorDescription(HRESULT_FROM_WIN32(GetLastError()));
		return EXIT_FAILURE;
	}

	hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE | COINIT_SPEED_OVER_MEMORY);
    if (FAILED(hr))
    {
		ErrorDescription(hr);
        return EXIT_FAILURE;
    }

	hr = OleInitialize(0); // Drag and drop
	if (FAILED(hr))
    {
		ErrorDescription(hr);
        return EXIT_FAILURE;
    }

	// Initialize global strings
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_MEP, szWindowClass, MAX_LOADSTRING);

    // Initialize device-indpendent resources, such as the Direct2D factory.
    hr = renderer.CreateDeviceIndependentResources();

    if (SUCCEEDED(hr))
    {
		MyRegisterClass(hInstance);
		MyRegisterRenderTargetClass(hInstance);
	}

	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow))
	{
		return EXIT_FAILURE;
	}

	//g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_ALLPROPERTIES, NULL);
	//g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_ALLPROPERTIES, NULL);

	g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_Label);
	g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
	g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_TooltipTitle);
	g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_TooltipDescription);
	g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_SmallImage);
	g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable1, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_LargeImage);
	
	g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_Label);
	g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
	g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_TooltipTitle);
	g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_TooltipDescription);
	g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_SmallImage);
	g_pRibbonFramework->InvalidateUICommand(cmdHome_EntropyGrid_InitialisationVariable2, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_LargeImage);

	hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_MEP));

	while (true)  // LOOP FOREVER
    {
        // 1.  PROCESS MESSAGES FROM WINDOWS
        if (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
        {
            // Now, this is DIFFERENT than GetMessage()!
            // PeekMessage is NOT a BLOCKING FUNCTION.
            // That is, PeekMessage() returns immediately
            // with either a true (there was a message
            // for our window)
            // or a false (there was no message for our window).

            // If there was a message for our window, then
            // we want to translate and dispatch that message.

            // otherwise, we want to be free to run
            // the next frame of our program.
            if (msg.message == WM_QUIT)
            {
                break;  // BREAK OUT OF INFINITE LOOP if user is trying to quit
            }
            
			if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
			{
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}
            // translates 
            // this line RESULTS IN
            // a call to WndProc(), passing the message and
            // the HWND.
            
            // ALL OTHER application processing happens
            // in the lines of code below
        }
		
		if (g_Variables.Running)
		{
			LARGE_INTEGER PerformanceCounterInit;
QueryPerformanceCounter(&PerformanceCounterInit);

			if (g_Variables.OutputToFile)
			{
				OutputToFile();
			}

			RunSimulation();

			CalculatePopulation();

			CalculateThermodynamicEntropy();

			LARGE_INTEGER PerformanceCounter;
QueryPerformanceCounter(&PerformanceCounter);
LARGE_INTEGER PerformanceFrequency;
QueryPerformanceFrequency(&PerformanceFrequency);

g_Variables.FPS = (1.0/(((double)PerformanceCounter.QuadPart - (double)PerformanceCounterInit.QuadPart)/(double)PerformanceFrequency.QuadPart));

			g_Variables.Redraw = true;

		}
		else
		{
			::Sleep(1);
			/*if (g_Variables.Generation != 0)
			{
				g_pTaskbarList->SetProgressState(hWnd, TBPF_PAUSED); //keeps looping
			}*/
		}

		if (!IsIconic(hWnd)) //if not minimized
		{
			UpdateStatusBar();

			if (g_Variables.Redraw && g_Variables.Visualisation) //minimal evaluation as most of the time Redraw will be false
			{
				renderer.OnRender();
			}
		}
    }

	myfile.close();
	CoUninitialize();
	OleUninitialize();

	return (int)msg.wParam;
}



//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
//  COMMENTS:
//
//    This function and its usage are only necessary if you want this code
//    to be compatible with Win32 systems prior to the 'RegisterClassEx'
//    function that was added to Windows 95. It is important to call this function
//    so that the application will get 'well formed' small icons associated
//    with it.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MEP));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= NULL; // (HBRUSH)GetStockObject(BLACK_BRUSH);//(HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= MAKEINTRESOURCE(IDC_MEP);
	wcex.lpszClassName	= szWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

	return RegisterClassEx(&wcex);
}

ATOM MyRegisterRenderTargetClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= /*CS_HREDRAW | CS_VREDRAW |*/ CS_PARENTDC;
	wcex.lpfnWndProc	= RenderTargetWndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= NULL;
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= NULL; // (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= NULL;
	wcex.lpszClassName	= TEXT("RenderTarget");
	wcex.hIconSm		= NULL;

	return RegisterClassEx(&wcex);
}

void InitialiseArrays()
{
	g_Variables.BitmapDataEntropy = new byte[g_Variables.Stride*g_Variables.Height];
	g_Variables.BitmapDataLife = new byte[g_Variables.Stride*g_Variables.Height];

	Concurrency::parallel_for(0, (int)(g_Variables.Stride*g_Variables.Height), [&](UINT i)
	{
		g_Variables.BitmapDataEntropy[i] = 0;
		g_Variables.BitmapDataLife[i] = 0;
	});


	g_Variables.EntropyCurrent = new UINT*[g_Variables.Height];
	g_Variables.EntropyNext = new UINT*[g_Variables.Height];
	
	Concurrency::parallel_for(0, (int)g_Variables.Height, [&](UINT i)
	{
		g_Variables.EntropyCurrent[i] = new UINT[g_Variables.Width];
		g_Variables.EntropyNext[i] = new UINT[g_Variables.Width];
	});
		
	Concurrency::parallel_for(0, (int)g_Variables.Height, [&](UINT i)
	{
		for (UINT j = 0; j < g_Variables.Width; j++)
		{
			g_Variables.EntropyCurrent[i][j] = 0;
			g_Variables.EntropyNext[i][j] = 0;
		}
	});
}

inline void CalculatePopulation()
{
	g_Variables.Population = (LONG)g_Variables.EcosystemCurrent.size();
}

inline void CalculateThermodynamicEntropy()
{
	/*UINT TotalFreeEnergy = 0;

	for (UINT i = 0; i < g_Variables.Height; i++)
	{
		for (UINT j = 0; j < g_Variables.Width; j++)
		{
			TotalFreeEnergy += g_Variables.EntropyCurrent[i][j];
		}
	};*/

	/*UINT TotalFreeEnergy = 0;
	Concurrency::combinable<int> sum;

	Concurrency::parallel_for(0, (int)g_Variables.Height, [&](UINT i)
	{
		for (UINT j = 0; j < g_Variables.Width; j++)
		{
			sum.local() += g_Variables.EntropyCurrent[i][j];
		}
	});

	TotalFreeEnergy = sum.combine(std::plus<int>());*/

	if (g_Variables.FreeEnergy == 0)
	{
		g_Variables.ThermodynamicEntropy = 0;
	}
	else
	{
		g_Variables.ThermodynamicEntropy = 1.0 - (double)g_Variables.FreeEnergy/(double)g_Variables.FreeEnergyInitial;
	}
}

void UpdateStatusBar()
{
	WCHAR wGeneration[MAX_LOADSTRING];
	WCHAR wPopulation[MAX_LOADSTRING];
	WCHAR wThermodynamicEntropy[MAX_LOADSTRING];
	WCHAR wFPS[MAX_LOADSTRING];

	HRESULT hr = StringCchPrintf(wGeneration, MAX_LOADSTRING, TEXT("Generation: %d"), g_Variables.Generation);
	if (SUCCEEDED(hr))
	{
		SendMessage(hStatusBar, SB_SETTEXT, (WPARAM) 1, (LPARAM) wGeneration);
	}
	
	hr = StringCchPrintf(wPopulation, MAX_LOADSTRING, L"Population: %d", g_Variables.Population);
	if (SUCCEEDED(hr))
	{
		SendMessage(hStatusBar, SB_SETTEXT, (WPARAM) 2, (LPARAM) wPopulation);
	}
	
	hr = StringCchPrintf(wThermodynamicEntropy, MAX_LOADSTRING, L"Entropy: %f", g_Variables.ThermodynamicEntropy);
	if (SUCCEEDED(hr))
	{
		SendMessage(hStatusBar, SB_SETTEXT, (WPARAM) 3, (LPARAM) wThermodynamicEntropy);
	}
	
	hr = StringCchPrintf(wFPS, MAX_LOADSTRING, L"FPS: %d", (UINT)g_Variables.FPS);
	if (SUCCEEDED(hr))
	{
		SendMessage(hStatusBar, SB_SETTEXT, (WPARAM) 4, (LPARAM) wFPS);
	}
}

void RunSimulation()
{
	UINT Counter = 0U;

	//g_Variables.EcosystemNext.clear();

	//Concurrency::critical_section cs;

	UINT CounterTotalOrganisms = (UINT)g_Variables.EcosystemCurrent.size();
	UINT CounterOrganisms = 0U;

	std::list<LIFEPROPERTIES>::iterator Iter = g_Variables.EcosystemCurrent.begin();

	while (Iter != g_Variables.EcosystemCurrent.end() && CounterOrganisms != CounterTotalOrganisms)
	{
		CounterOrganisms++;

		//UINT Neighbourhood[10] = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0};

		UINT i, j, NextDirectioni, NextDirectionj;
		UINT BestValue;

		j = Iter->x;
		i = Iter->y;

		/*if (Iter->Lifespan == 0)
		{
			Iter = g_Variables.EcosystemCurrent.erase(Iter);
			continue;
		}*/

		if (j == 0 || j == g_Variables.Width - 1 || i == 0 || i == g_Variables.Height - 1)
		{
			Iter = g_Variables.EcosystemCurrent.erase(Iter);
			continue;
		}

		NextDirectioni = i;
		NextDirectionj = j;

		BestValue = g_Variables.EntropyCurrent[i][j];

		if (BestValue == 0)
		{
			Iter = g_Variables.EcosystemCurrent.erase(Iter);
			continue;
		}

		if (g_Variables.DeathRate != 0)
		{
			if (RandomNumber::Double(g_Variables.RandomNumberFunction) > (1.0 - g_Variables.DeathRate))
			{
				Iter = g_Variables.EcosystemCurrent.erase(Iter);
				continue;
			}
		}

		if (g_Variables.ReproductionRate != 0)
		{
			if (RandomNumber::Double(g_Variables.RandomNumberFunction) > (1.0 - g_Variables.ReproductionRate))
			{
				LIFEPROPERTIES lpReproductionRate;
				
				lpReproductionRate.x = j;
				lpReproductionRate.y = i;
				lpReproductionRate.sex = (SEX)RandomNumber::Uniform(g_Variables.RandomNumberFunction, 0, 2);
				lpReproductionRate.Lifespan = 100;

				g_Variables.EcosystemCurrent.push_back(lpReproductionRate);
				//g_Variables.EcosystemNext.push_back(lpReproductionRate);
			}
		}

		if (g_Variables.Intelligence)
		{
			if ((g_Variables.Generation%2 == 0) || g_Variables.Bias)
			{
				//Top
				if (g_Variables.EntropyCurrent[i-1][j] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i-1][j];
					NextDirectioni = i-1;
					NextDirectionj = j;
				}
				//TopRight
				if (g_Variables.EntropyCurrent[i-1][j+1] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i-1][j+1];
					NextDirectioni = i-1;
					NextDirectionj = j+1;
				}
				//Right
				if (g_Variables.EntropyCurrent[i][j+1] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i][j+1];
					NextDirectioni = i;
					NextDirectionj = j+1;
				}
				//BottomRight
				if (g_Variables.EntropyCurrent[i+1][j+1] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i+1][j+1];
					NextDirectioni = i+1;
					NextDirectionj = j+1;
				}
				//Bottom
				if (g_Variables.EntropyCurrent[i+1][j] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i+1][j];
					NextDirectioni = i+1;
					NextDirectionj = j;
				}
				//BottomLeft
				if (g_Variables.EntropyCurrent[i+1][j-1] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i+1][j-1];
					NextDirectioni = i+1;
					NextDirectionj = j-1;
				}
				//Left
				if (g_Variables.EntropyCurrent[i][j-1] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i][j-1];
					NextDirectioni = i;
					NextDirectionj = j-1;
				}
				//TopLeft
				if (g_Variables.EntropyCurrent[i-1][j-1] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i-1][j-1];
					NextDirectioni = i-1;
					NextDirectionj = j-1;
				}
			}
			else
			{
				//Bottom
				if (g_Variables.EntropyCurrent[i+1][j] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i+1][j];
					NextDirectioni = i+1;
					NextDirectionj = j;
				}
				//BottomLeft
				if (g_Variables.EntropyCurrent[i+1][j-1] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i+1][j-1];
					NextDirectioni = i+1;
					NextDirectionj = j-1;
				}
				//Left
				if (g_Variables.EntropyCurrent[i][j-1] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i][j-1];
					NextDirectioni = i;
					NextDirectionj = j-1;
				}
				//TopLeft
				if (g_Variables.EntropyCurrent[i-1][j-1] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i-1][j-1];
					NextDirectioni = i-1;
					NextDirectionj = j-1;
				}
				//Top
				if (g_Variables.EntropyCurrent[i-1][j] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i-1][j];
					NextDirectioni = i-1;
					NextDirectionj = j;
				}
				//TopRight
				if (g_Variables.EntropyCurrent[i-1][j+1] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i-1][j+1];
					NextDirectioni = i-1;
					NextDirectionj = j+1;
				}
				//Right
				if (g_Variables.EntropyCurrent[i][j+1] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i][j+1];
					NextDirectioni = i;
					NextDirectionj = j+1;
				}
				//BottomRight
				if (g_Variables.EntropyCurrent[i+1][j+1] > BestValue)
				{
					BestValue = g_Variables.EntropyCurrent[i+1][j+1];
					NextDirectioni = i+1;
					NextDirectionj = j+1;
				}
			}		
		} //Intelligent
		else //if not intelligent
		{
			UINT choice = RandomNumber::Uniform(g_Variables.RandomNumberFunction, 0, 8);

			if((g_Variables.Generation % 2 == 0) || g_Variables.Bias)
			{
				if(choice == 0) //Top
				{
					NextDirectioni = i-1;
					NextDirectionj = j;
				}
				if(choice == 1) //TopRight
				{
					NextDirectioni = i-1;
					NextDirectionj = j+1;
				}
				if(choice == 2) //Right
				{
					NextDirectioni = i;
					NextDirectionj = j+1;
				}
				if(choice == 3) //BottomRight
				{
					NextDirectioni = i+1;
					NextDirectionj = j+1;
				}
				if(choice == 4) //Bottom
				{
					NextDirectioni = i+1;
					NextDirectionj = j;							
				}
				if(choice == 5) //BottomLeft
				{
					NextDirectioni = i+1;
					NextDirectionj = j-1;
				}
				if(choice == 6) //Left
				{
					NextDirectioni = i;
					NextDirectionj = j-1;
				}
				if(choice == 7) //TopLeft
				{
					NextDirectioni = i-1;
					NextDirectionj = j-1;
				}
			}
			else
			{
				if(choice == 0) //Bottom
				{
					NextDirectioni = i+1;
					NextDirectionj = j;							
				}
				if(choice == 1) //BottomLeft
				{
					NextDirectioni = i+1;
					NextDirectionj = j-1;
				}
				if(choice == 2) //Left
				{
					NextDirectioni = i;
					NextDirectionj = j-1;
				}
				if(choice == 3) //TopLeft
				{
					NextDirectioni = i-1;
					NextDirectionj = j-1;
				}
				if(choice == 4) //Top
				{
					NextDirectioni = i-1;
					NextDirectionj = j;
				}
				if(choice == 5) //TopRight
				{
					NextDirectioni = i-1;
					NextDirectionj = j+1;
				}
				if(choice == 6) //Right
				{
					NextDirectioni = i;
					NextDirectionj = j+1;
				}
				if(choice == 7) //BottomRight
				{
					NextDirectioni = i+1;
					NextDirectionj = j+1;
				}
			}
		}

		//LIFEPROPERTIES lpNext;

		//lpNext.x = NextDirectionj;
		//lpNext.y = NextDirectioni;

		//g_Variables.EcosystemNext.push_back(lpNext);

		Iter->x = NextDirectionj;
		Iter->y = NextDirectioni;
		Iter->Lifespan--;

		//InterlockedDecrement(&PopulationValueArray[current[i][j]]);
		//InterlockedIncrement(&PopulationValueArray[(current[i][j] - 1)]);
		//cs.lock();
		//InterlockedDecrement(&g_Variables.EntropyCurrent[i][j]);
		//InterlockedDecrement(&g_Variables.FreeEnergy);
		g_Variables.EntropyCurrent[i][j]--;
		g_Variables.FreeEnergy--;
		//cs.unlock();
		Iter++;
   };

			
	if (g_Variables.Dissipation == DISSIPATION_AVERAGECEIL)
	{
		for (UINT i = 0; i < g_Variables.Height; i++)
		{
			g_Variables.EntropyNext[i][0] = 0;
			g_Variables.EntropyNext[i][g_Variables.Width-1] = 0;
		}

		for (UINT j = 0; j < g_Variables.Width; j++)
		{
			g_Variables.EntropyNext[0][j] = 0;
			g_Variables.EntropyNext[g_Variables.Height-1][j] = 0;
		}

		Concurrency::parallel_for(1, (int)g_Variables.Height - 1, [&](UINT i)
		{
			for (UINT j = 1; j < g_Variables.Width - 1; j++)
			{
				UINT NextValue;
			
				NextValue = (UINT)ceil(
				(double)(
				g_Variables.EntropyCurrent[i][j]		+	//Current
				g_Variables.EntropyCurrent[i-1][j]		+	//Top
				g_Variables.EntropyCurrent[i+1][j]		+	//Bottom
				g_Variables.EntropyCurrent[i][j-1]		+	//Left
				g_Variables.EntropyCurrent[i][j+1]		+	//Right
				g_Variables.EntropyCurrent[i-1][j-1]	+	//TopLeft
				g_Variables.EntropyCurrent[i-1][j+1]	+	//TopRight
				g_Variables.EntropyCurrent[i+1][j-1]	+	//BottomLeft
				g_Variables.EntropyCurrent[i+1][j+1]		//BottomRight
				)/9.0);

				//InterlockedDecrement(&PopulationValueArray[current[i][j]]);
				//InterlockedIncrement(&PopulationValueArray[NextValue]);

				g_Variables.EntropyNext[i][j] = NextValue;
			}
		});

		//Copy next to current
		Concurrency::parallel_for(0, (int)g_Variables.Height, [&](UINT i)
		{
			for(UINT j = 0; j < g_Variables.Width; j++)
			{
				g_Variables.EntropyCurrent[i][j] = g_Variables.EntropyNext[i][j];
			}
		});
	}

	if (g_Variables.Dissipation == DISSIPATION_AVERAGEFLOOR)
	{
		for (UINT i = 0; i < g_Variables.Height; i++)
		{
			g_Variables.EntropyNext[i][0] = 0;
			g_Variables.EntropyNext[i][g_Variables.Width-1] = 0;
		}

		for (UINT j = 0; j < g_Variables.Width; j++)
		{
			g_Variables.EntropyNext[0][j] = 0;
			g_Variables.EntropyNext[g_Variables.Height-1][j] = 0;
		}

		Concurrency::parallel_for(1, (int)g_Variables.Height - 1, [&](UINT i)
		{
			for (UINT j = 1; j < g_Variables.Width - 1; j++)
			{
				UINT NextValue;
			
				NextValue =
				(
				g_Variables.EntropyCurrent[i][j]		+	//Current
				g_Variables.EntropyCurrent[i-1][j]		+	//Top
				g_Variables.EntropyCurrent[i+1][j]		+	//Bottom
				g_Variables.EntropyCurrent[i][j-1]		+	//Left
				g_Variables.EntropyCurrent[i][j+1]		+	//Right
				g_Variables.EntropyCurrent[i-1][j-1]	+	//TopLeft
				g_Variables.EntropyCurrent[i-1][j+1]	+	//TopRight
				g_Variables.EntropyCurrent[i+1][j-1]	+	//BottomLeft
				g_Variables.EntropyCurrent[i+1][j+1]		//BottomRight
				)/9;			

				//InterlockedDecrement(&PopulationValueArray[current[i][j]]);
				//InterlockedIncrement(&PopulationValueArray[NextValue]);

				g_Variables.EntropyNext[i][j] = NextValue;
			}
		});
	
		//Copy next to current
		Concurrency::parallel_for(0, (int)g_Variables.Height, [&](UINT i)
		{
			for(UINT j = 0; j < g_Variables.Width; j++)
			{
				g_Variables.EntropyCurrent[i][j] = g_Variables.EntropyNext[i][j];
			}
		});
	}

	g_Variables.Generation++;

	Counter++;

	g_pTaskbarList->SetProgressState(hWnd, TBPF_NORMAL);
	g_pTaskbarList->SetProgressValue(hWnd, (ULONGLONG)(g_Variables.ThermodynamicEntropy*1000000), 1000000);

	if ((Counter == g_Variables.NumberOfSteps) || (g_Variables.Population == 0))
	{
		g_Variables.Running = false;
		g_Variables.NumberOfSteps = 0;

		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_BooleanValue);
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_Label);
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_LargeImage);
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_SmallImage);

		if (g_Variables.Population == 0)
		{
			FLASHWINFO fwi;

			fwi.cbSize = sizeof(FLASHWINFO);
			fwi.dwFlags = FLASHW_ALL;
			fwi.dwTimeout = 0;
			fwi.hwnd = hWnd;
			fwi.uCount = 1U;

			FlashWindowEx(&fwi);

			g_pTaskbarList->SetProgressState(hWnd, TBPF_ERROR);
		}
	}

} //RunSimulation

void InitialiseVariables()
{
	g_Variables.NewRun = true;
	g_Variables.BitsPerPixel = 32U;
	g_Variables.BytesPerPixel = g_Variables.BitsPerPixel/8U;
	g_Variables.Stride = g_Variables.Width * g_Variables.BytesPerPixel;
	g_Variables.Visualisation = true;
	g_Variables.OutputToFile = false;
	g_Variables.Paint = false;
	g_Variables.Redraw = true;
	g_Variables.Running = false;
	g_Variables.NumberOfSteps = 0U;
	g_Variables.RandomNumberFunction = RAN;
	g_Variables.PeriodicBoundaryConditions = false;
	g_Variables.Initialisation = INITIALISATION_NORMAL;
	g_Variables.UniformMin = 0U;
	g_Variables.UniformMax = (UINT)pow(2.0, (int)((g_Variables.BitsPerPixel - 8)/3)); //-8 for unused Alpha channel
	g_Variables.NormalMean = (UINT)pow(2.0, (int)((g_Variables.BitsPerPixel - 8)/3))/2 - 1; //-8 for unused Alpha channel
	g_Variables.NormalVariance = 20;
	g_Variables.Dissipation = DISSIPATION_NONE;
	g_Variables.PopulationDefault = (LONG)((g_Variables.Height*g_Variables.Width)/262.144);
	g_Variables.PopulationInitial = g_Variables.PopulationDefault;
	g_Variables.Population = 0L;
	g_Variables.ReproductionRate = 0;
	g_Variables.DeathRate = 0;
	g_Variables.Intelligence = true;
	g_Variables.Bias = false;
	g_Variables.PaintLayer = LAYER_LIFE;
	g_Variables.Generation = 0u;
	g_Variables.InformationEntropy = 0;
	g_Variables.FreeEnergyInitial = 0u;
	g_Variables.FreeEnergy = 0u;
	g_Variables.ThermodynamicEntropy = 0;
	g_Variables.EntropyInitialised = false;
	g_Variables.LifeInitialised = false;
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
bool InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	hInst = hInstance; // Store instance handle in our global variable

	InitialiseVariables();
	InitialiseArrays();

	hWnd = CreateWindowExW(
		0,
		szWindowClass,
		szTitle,
		WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,//static_cast<UINT>(ceil(1280.f * dpiX / 96.f)),
		0,//static_cast<UINT>(ceil(480.f * dpiY / 96.f)),
		NULL,
		NULL,
		hInstance,
		NULL);

	if (!hWnd)
	{
		return false;
	}

	ShowWindow(hWnd, nCmdShow);
	
	UpdateWindow(hWnd);

	RECT rc = {0};
	GetClientRect(hStatusBar, &rc);

	ClientResize(hWnd, g_Variables.Width, (g_Variables.Height + GetRibbonHeight() + (rc.bottom - rc.top)));

	return true;
}

POINT RenderTargetTranslate(HWND hWnd, UINT width, UINT height)
{
	POINT point = {0};
	RECT rc = {0};

	GetClientRect(hWnd, &rc);

	point.x = ((rc.right - rc.left)/2) - (g_metrics.ScaleX(width)/2);
	point.y = ((rc.bottom - rc.top)/2) - (g_metrics.ScaleY(height)/2);

	point.x = g_metrics.UnscaleX(point.x);
	point.y = g_metrics.UnscaleY(point.y);

	if (point.x < 0)
	{
		point.x = 0;
	}

	if (point.y < 0)
	{
		point.y = 0;
	}

	return point;
}

LRESULT CALLBACK RenderTargetWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	static bool mouseTracked = false;

	switch (message)
	{

	case WM_ERASEBKGND: //To prevent flicker
		{
		}
		return 1;

	case WM_DISPLAYCHANGE:
	case WM_PAINT:
		{
			HDC hdc;
			PAINTSTRUCT ps;

			hdc = BeginPaint(hRenderTarget, &ps);

			renderer.OnRender();

			EndPaint(hRenderTarget, &ps);
		}
		return 0;

	case WM_SIZE:
		{
			UINT width = LOWORD(lParam);
            UINT height = HIWORD(lParam);

			g_Variables.TranslateVector = RenderTargetTranslate(hWnd, g_Variables.Width, g_Variables.Height);

            renderer.OnResize(width, height);
		}
		return 0;

	/*case WM_MOUSEHOVER:
		{
			UINT x = GET_X_LPARAM(lParam);
			UINT y = GET_Y_LPARAM(lParam);

			HRESULT hr;
			TCHAR wa[MAX_LOADSTRING];

			hr = StringCchPrintf(wa, MAX_LOADSTRING, TEXT("%d, %dpx"), x, y);
			if (SUCCEEDED(hr))
			{
				MessageBox(NULL, wa, TEXT("Cursor"), MB_OK | MB_TASKMODAL);
			}
		}
		return 0;*/

	case WM_LBUTTONDOWN:
		{
			if (g_Variables.Paint)
			{
				UINT x = GET_X_LPARAM(lParam);
				UINT y = GET_Y_LPARAM(lParam);

				if ((x >= (UINT)g_Variables.TranslateVector.x) && (y >= (UINT)g_Variables.TranslateVector.y) && (x < g_Variables.TranslateVector.x + g_Variables.Width) && (y < g_Variables.TranslateVector.y + g_Variables.Height))
				{
					SetCapture(hWnd);

					if (g_Variables.PaintLayer == LAYER_LIFE)
					{
						LIFEPROPERTIES lp;

						lp.x = x - g_Variables.TranslateVector.x;
						lp.y = y - g_Variables.TranslateVector.y;

						g_Variables.EcosystemCurrent.push_back(lp);
						
						//UINT index = (lp.x * (g_Variables.BytesPerPixel)) + (lp.y * g_Variables.Stride);

						//g_Variables.BitmapDataLife[(index + 2)] = (byte)(255); //Red
						//g_Variables.BitmapDataLife[(index + 3)] = (byte)(255); //Alpha

						g_Variables.Population++;

						g_Variables.Redraw = true;
						//renderer.OnRender();
						UpdateStatusBar();

						g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
						g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_NextStep, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
						g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Clear, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
						g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearEntropyGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
						g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearLifeGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
						g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearSplitButton, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
					}
				}
			}
		}
		return 0;

	case WM_LBUTTONUP:
		{
			mouseTracked = false;
			ReleaseCapture();
		}
		return 0;

	/*case WM_RBUTTONDOWN:
		{
			if (g_Variables.Paint)
			{
				UINT x = GET_X_LPARAM(lParam);
				UINT y = GET_Y_LPARAM(lParam);

				if ((x >= (UINT)g_Variables.TranslateVector.x) && (y >= (UINT)g_Variables.TranslateVector.y) && (x < g_Variables.TranslateVector.x + g_Variables.Width) && (y < g_Variables.TranslateVector.y + g_Variables.Height))
				{
					SetCapture(hWnd);

					if (g_Variables.PaintLayer == LAYER_LIFE)
					{
						if (g_Variables.LifeCurrent[y - g_Variables.TranslateVector.y][x - g_Variables.TranslateVector.x] != 0)
						{
							g_Variables.LifeCurrent[y - g_Variables.TranslateVector.y][x - g_Variables.TranslateVector.x] = 0;
							
							g_Variables.Population--;

							g_Variables.Redraw = true;
							UpdateStatusBar();
						}

						g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
						g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_NextStep, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
						g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Clear, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
						g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearEntropyGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
						g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearLifeGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
						g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearSplitButton, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
					}
				}
			}
		}
		return 0;*/

	case WM_RBUTTONUP:
		{
			mouseTracked = false;
			ReleaseCapture();
		}
		return 0;

	case WM_MOUSELEAVE:
		{
			mouseTracked = false;
			SendMessage(hStatusBar, SB_SETTEXT, (WPARAM) 0, (LPARAM) L"");
			//SendMessage(hStatusBar, SB_SETTEXT, (WPARAM) 4, (LPARAM) L"");
		}
		return 0;

	case WM_MOUSEMOVE:
		{
			UINT x = GET_X_LPARAM(lParam);
			UINT y = GET_Y_LPARAM(lParam);

			x = g_metrics.UnscaleX(x);
			y = g_metrics.UnscaleY(y);

			if (!mouseTracked)
			{
				TRACKMOUSEEVENT tme;
				tme.cbSize = sizeof(TRACKMOUSEEVENT);
				tme.dwFlags = /*TME_HOVER |*/ TME_LEAVE;
				tme.dwHoverTime = 500;
				tme.hwndTrack = hWnd;
				if (TrackMouseEvent(&tme))
				{
					mouseTracked = true;
				}
				else
				{
					ErrorDescription(HRESULT_FROM_WIN32(GetLastError()));
				}
			}

			if ((x >= (UINT)g_Variables.TranslateVector.x) && (y >= (UINT)g_Variables.TranslateVector.y) && (x < g_Variables.TranslateVector.x + g_Variables.Width) && (y < g_Variables.TranslateVector.y + g_Variables.Height))
			{
				WCHAR wMouseCoordinates[MAX_LOADSTRING];

				if (g_Variables.EntropyInitialised)
				{
					StringCchPrintf(wMouseCoordinates, MAX_LOADSTRING, TEXT("%u, %upx (%u)"), x - g_Variables.TranslateVector.x, y - g_Variables.TranslateVector.y, g_Variables.EntropyCurrent[y - g_Variables.TranslateVector.y][x - g_Variables.TranslateVector.x]);
				}
				else
				{
					StringCchPrintf(wMouseCoordinates, MAX_LOADSTRING, TEXT("%u, %upx"), x - g_Variables.TranslateVector.x, y - g_Variables.TranslateVector.y);
				}

				SendMessage(hStatusBar, SB_SETTEXT, (WPARAM) 0, (LPARAM) wMouseCoordinates);

				if (g_Variables.Paint)
				{
					if (g_Variables.PaintLayer == LAYER_LIFE)
					{
						if (wParam & MK_LBUTTON)
						{
							LIFEPROPERTIES lp;

							lp.x = x - g_Variables.TranslateVector.x;
							lp.y = y - g_Variables.TranslateVector.y;

							g_Variables.EcosystemCurrent.push_back(lp);
							
							//UINT index = (lp.x * (g_Variables.BytesPerPixel)) + (lp.y * g_Variables.Stride);

							//g_Variables.BitmapDataLife[(index + 2)] = (byte)(255); //Red
							//g_Variables.BitmapDataLife[(index + 3)] = (byte)(255); //Alpha

							g_Variables.Population++;

							g_Variables.Redraw = true;
							//renderer.OnRender();
							UpdateStatusBar();
						}
						/*else if (wParam & MK_RBUTTON)
						{
							if (g_Variables.LifeCurrent[y - translateVector.y][x - translateVector.x] != 0)
							{
								g_Variables.LifeCurrent[y - translateVector.y][x - translateVector.x] = 0;
							
								g_Variables.Population--;

								g_Variables.Redraw = true;
								UpdateStatusBar();
							}
						}*/
					}

					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_NextStep, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Clear, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearEntropyGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearLifeGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearSplitButton, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				}
			}
			else
			{
				SendMessage(hStatusBar, SB_SETTEXT, (WPARAM) 0, (LPARAM) L"");
				//SendMessage(hStatusBar, SB_SETTEXT, (WPARAM) 4, (LPARAM) L"");
			}
		}
		return 0;

	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

void ClientResize(HWND &hWnd, UINT nWidth, UINT nHeight) 
{
	RECT rcClient, rcWind = {0};
	POINT ptDiff = {0};
	
	GetClientRect(hWnd, &rcClient);
	GetWindowRect(hWnd, &rcWind);
	ptDiff.x = (rcWind.right - rcWind.left) - rcClient.right;
	ptDiff.y = (rcWind.bottom - rcWind.top) - rcClient.bottom;
	
	MoveWindow(hWnd, rcWind.left, rcWind.top, nWidth + ptDiff.x, nHeight + ptDiff.y, false);
}

void CreateLayerFromFile(PCWSTR FileName, LAYER layer)
{
	//SetWindowText(hWnd, FileName);

	HRESULT hr = renderer.CreateLayerFromFile(FileName, layer);
	if (FAILED(hr))
	{
		//ErrorDescription(hr);
		MessageBox(hWnd, L"The image could not be loaded.", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
		return;
	}

	if (layer == LAYER_ENTROPY)
	{
		g_Variables.FreeEnergyInitial = 0;

		for (UINT i = 0; i < g_Variables.Height; i++)
		{
			for (UINT j = 0; j < g_Variables.Width; j++) 
			{
				UINT CurrentToBeWritten;
							
				CurrentToBeWritten = (UINT)g_Variables.BitmapDataEntropy[((i*g_Variables.Width) + j)*g_Variables.BytesPerPixel];

				g_Variables.EntropyCurrent[i][j] = CurrentToBeWritten;

				g_Variables.FreeEnergyInitial = g_Variables.FreeEnergyInitial + CurrentToBeWritten;
			}
		}

		g_Variables.FreeEnergy = g_Variables.FreeEnergyInitial;

		CalculateThermodynamicEntropy();

		g_Variables.EntropyInitialised = true;
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_NextStep, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Clear, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearEntropyGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearLifeGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearSplitButton, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
	}
	else if (layer == LAYER_LIFE)
	{
		g_Variables.EcosystemCurrent.clear();

		g_Variables.Population = 0;

		for (UINT i = 0; i < g_Variables.Height; i++)
		{
			for (UINT j = 0; j < g_Variables.Width; j++) 
			{
				if ((UINT)g_Variables.BitmapDataLife[(((i*g_Variables.Width) + j)*g_Variables.BytesPerPixel)+2] == 255) //+2 to get red channel
				{
					LIFEPROPERTIES lp;
				
					lp.x = j;
					lp.y = i;

					g_Variables.EcosystemCurrent.push_back(lp);
					g_Variables.Population++;
				}
			}
		}

		g_Variables.LifeInitialised = true;
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_NextStep, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Clear, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearEntropyGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearLifeGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
		g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearSplitButton, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
	}

	g_Variables.Redraw = true;
	UpdateStatusBar();
}

HRESULT CommonItemDialogOpen(HWND hWnd, wchar_t *buffer)
{
	// Initialize COM engine
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE); // Will return S_FALSE (1) if already initialised
 
    if (SUCCEEDED(hr))
    {
		IFileDialog *pfd = nullptr;
		// CoCreate the dialog object.
		hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pfd));

		if (SUCCEEDED(hr))
		{
			COMDLG_FILTERSPEC rgSpec[] =
				{
					{ L"All image files", L"*.bmp;*.dib;*.gif;*.jpg;*.jpeg;*.jpe;*.jif;*.jfif;*.jfi;*.png;*.tif;*.tiff" },
					{ L"BMP", L"*.bmp;*.dib" },
					{ L"GIF", L"*.gif" },
					{ L"JPG", L"*.jpg;*.jpeg;*.jpe;*.jif;*.jfif;*.jfi" },
					{ L"PNG", L"*.png" },
					{ L"TIFF", L"*.tif;*.tiff" },
					{ L"All", L"*.*" }
				};

			hr = pfd->SetFileTypes(sizeof(rgSpec)/sizeof(*rgSpec), rgSpec);

			if (SUCCEEDED(hr))
			{
				// Show the dialog
				hr = pfd->Show(hWnd);
        
				if (SUCCEEDED(hr))
				{
					IShellItem *psiResult;
					// Obtain the result of the user's interaction with the dialog.
					hr = pfd->GetResult(&psiResult);
            
					if (SUCCEEDED(hr))
					{
						LPWSTR pszFileSysPath;

						hr = psiResult->GetDisplayName(SIGDN_FILESYSPATH, &pszFileSysPath);
						if (SUCCEEDED(hr))
						{
							hr = StringCchCopyW(buffer, MAX_PATH, pszFileSysPath);
							CoTaskMemFree(pszFileSysPath);
						}
						psiResult->Release();
					}
				}
			}
			pfd->Release();
		}
	}

	// Cleanup COM
    CoUninitialize(); // Must be called for each CoInitialize/CoInitializeEx

	return hr;
}

void CreateFileFromLayer(PCWSTR FileName, LAYER layer, GUID guidContainerFormat)
{
	//SetWindowText(hWnd, FileName);

	HRESULT hr = renderer.CreateFileFromLayer(FileName, layer, guidContainerFormat);
	if (FAILED(hr))
	{
		ErrorDescription(hr);
		MessageBox(NULL, L"CreateFileFromLayer", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
	}
}

HRESULT CommonItemDialogSave(HWND hWnd, wchar_t *buffer, GUID *guidContainerFormat)
{
	// Initialize COM engine
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE); // Will return S_FALSE (1) if already initialised
 
    if (SUCCEEDED(hr))
    {
		IFileDialog *pfd = nullptr;
		// CoCreate the dialog object.
		hr = CoCreateInstance(CLSID_FileSaveDialog, 
										NULL, 
										CLSCTX_INPROC_SERVER, 
										IID_PPV_ARGS(&pfd));

		if (SUCCEEDED(hr))
		{
			COMDLG_FILTERSPEC rgSpec[] =
				{
					{ L"BMP", L"*.bmp" },
					{ L"PNG", L"*.png" },
					{ L"TIFF", L"*.tif" }
				};

			hr = pfd->SetFileTypes(sizeof(rgSpec)/sizeof(*rgSpec), rgSpec);

			if (SUCCEEDED(hr))
			{
				// Set default file type to png
				hr = pfd->SetFileTypeIndex(2);

				if (SUCCEEDED(hr))
				{
					hr = pfd->SetFileName(L"Untitled");

					if (SUCCEEDED(hr))
					{
					// Set default extension to png. This will update the default extension automatically when the user chooses a new file type
					hr = pfd->SetDefaultExtension(L"png");

						if (SUCCEEDED(hr))
						{
							// Show the dialog
							hr = pfd->Show(hWnd);

        					if (SUCCEEDED(hr))
							{
								// Obtain the result of the user's interaction with the dialog.
								IShellItem *psiResult;
								hr = pfd->GetResult(&psiResult);

								if (SUCCEEDED(hr))
								{
									LPWSTR pszFileSysPath;

									hr = psiResult->GetDisplayName(SIGDN_FILESYSPATH, &pszFileSysPath);
									if (SUCCEEDED(hr))
									{
										hr = StringCchCopyW(buffer, MAX_PATH, pszFileSysPath);
										if (SUCCEEDED(hr))
										{
											UINT fileTypeIndex;

											hr = pfd->GetFileTypeIndex(&fileTypeIndex);
											if (SUCCEEDED(hr))
											{
												if (fileTypeIndex == 1)
												{
													*guidContainerFormat = GUID_ContainerFormatBmp;
												}
												else if (fileTypeIndex == 2)
												{
													*guidContainerFormat = GUID_ContainerFormatPng;
												}
												else if (fileTypeIndex == 3)
												{
													*guidContainerFormat = GUID_ContainerFormatTiff;
												}
											}
										}
										CoTaskMemFree(pszFileSysPath);
									}
									psiResult->Release();
								}
							}
						}
					}
				}
			}
			pfd->Release();
		}
	}
	// Cleanup COM
    CoUninitialize(); // Must be called for each CoInitialize/CoInitializeEx

	return hr;
}

HRESULT CALLBACK TaskDialogCallbackProc(
  __in  HWND hWnd,
  __in  UINT uNotification,
  __in  WPARAM /*wParam*/,
  __in  LPARAM lParam,
  __in  LONG_PTR /*dwRefData*/
)
{
	if (uNotification == TDN_HYPERLINK_CLICKED)
	{
		ShellExecuteW(hWnd, L"open", (LPCWSTR)lParam, NULL, NULL, SW_SHOWNORMAL);
	}
	return S_OK;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	HDC hdc;
	PAINTSTRUCT ps;
	static IDropTarget *pDropTarget;
	static UINT s_uTBBC = WM_NULL;

    if (s_uTBBC == WM_NULL)
    {
        // Compute the value for the TaskbarButtonCreated message
        s_uTBBC = RegisterWindowMessageW(L"TaskbarButtonCreated");

        // In case the application is run elevated, allow the
        // TaskbarButtonCreated message through.
        ChangeWindowMessageFilter(s_uTBBC, MSGFLT_ADD);
    }

    if (message == s_uTBBC)
    {
        // Once we get the TaskbarButtonCreated message, we can call methods
        // specific to our window on a TaskbarList instance. Note that it's
        // possible this message can be received multiple times over the lifetime
        // of this window (if explorer terminates and restarts, for example).
        if (!g_pTaskbarList)
        {
            HRESULT hr = CoCreateInstance(CLSID_TaskbarList, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&g_pTaskbarList));
            if (SUCCEEDED(hr))
            {
                hr = g_pTaskbarList->HrInit();
                if (FAILED(hr))
                {
                    g_pTaskbarList->Release();
                    g_pTaskbarList = NULL;
                }
            }
        }
    }
	else switch (message)
	{
	case WM_COMMAND:
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam);
		
		switch (wmId)
		{
		case ID_FILE_NEW:
			{
				SendMessage(hWnd, WM_COMMAND, cmdHome_Commands_Clear, NULL);
			}
			break;
		case ID_FILE_OPEN:
			{
				MessageBoxW(hWnd, L"ID_FILE_OPEN", L"Info", MB_OK);
			}
			break;
		case ID_FILE_OPEN_ENTROPYGRID:
			{
				WCHAR buffer[MAX_PATH] = {0};

				HRESULT hr = CommonItemDialogOpen(hWnd, buffer);
				if (SUCCEEDED(hr))
				{
					CreateLayerFromFile(buffer, LAYER_ENTROPY);
				}
			}
			break;
		case ID_FILE_OPEN_LIFEGRID:
			{
				WCHAR buffer[MAX_PATH] = {0};

				HRESULT hr = CommonItemDialogOpen(hWnd, buffer);
				if (SUCCEEDED(hr))
				{
					CreateLayerFromFile(buffer, LAYER_LIFE);
				}
			}
			break;
		case ID_FILE_SAVE:
			{
				MessageBoxW(hWnd, L"ID_FILE_SAVE", L"Info", MB_OK);
			}
			break;
		case ID_FILE_SAVE_ENTROPYGRID:
			{
				WCHAR buffer[MAX_PATH] = {0};
				GUID guidContainerFormat = GUID_NULL;

				HRESULT hr = CommonItemDialogSave(hWnd, buffer, &guidContainerFormat);
				if (SUCCEEDED(hr))
				{
					CreateFileFromLayer(buffer, LAYER_ENTROPY, guidContainerFormat);
				}
			}
			break;
		case ID_FILE_SAVE_LIFEGRID:
			{
				WCHAR buffer[MAX_PATH] = {0};
				GUID guidContainerFormat = GUID_NULL;

				HRESULT hr = CommonItemDialogSave(hWnd, buffer, &guidContainerFormat);
				if (SUCCEEDED(hr))
				{
					CreateFileFromLayer(buffer, LAYER_LIFE, guidContainerFormat);
				}
			}
			break;

		case IDM_ABOUT:
		case ID_FILE_INFO:
		case ID_HELP:
			{
				TASKDIALOGCONFIG config		= {0};

				config.cbSize				= sizeof(config);
				config.dwCommonButtons		= TDCBF_CLOSE_BUTTON;
				config.dwFlags				= TDF_ALLOW_DIALOG_CANCELLATION | TDF_ENABLE_HYPERLINKS | TDF_POSITION_RELATIVE_TO_WINDOW;
				config.hInstance			= hInst;
				config.hwndParent			= hWnd;
				config.nDefaultButton		= IDCLOSE;
				config.pfCallback			= TaskDialogCallbackProc;
				config.pszContent			= L"Version 0.9\n"
											  L"<A HREF=\"http:\\maksymshostak.com\">Maksym Shostak</A>\n"
											  L"© 2011";
				config.pszMainIcon			= MAKEINTRESOURCEW(IDI_MEP); // TD_INFORMATION_ICON;
				config.pszMainInstruction	= L"MEP";
				config.pszWindowTitle		= L"About";
	
				HRESULT hr = TaskDialogIndirect(&config, NULL, NULL, NULL);
				if (FAILED(hr))
				{
					ErrorDescription(hr);
				}
			}
			break;

		case ID_FILE_EXIT:
			{
				DestroyWindow(hWnd);
			}
			break;

		case cmdHome_Options_Visualisation:
			{
				g_Variables.Visualisation = !g_Variables.Visualisation;

				g_pRibbonFramework->InvalidateUICommand(cmdHome_Options_Visualisation, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_BooleanValue);
			}
            break;

		case cmdHome_Options_OutputToFile:
			{
				g_Variables.OutputToFile = !g_Variables.OutputToFile;

				g_pRibbonFramework->InvalidateUICommand(cmdHome_Options_OutputToFile, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_BooleanValue);
			}
            break;

		case cmdHome_Options_Paint:
			{
				g_Variables.Paint = !g_Variables.Paint;

				g_pRibbonFramework->InvalidateUICommand(cmdHome_Options_Paint, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_BooleanValue);
				// Invalidate the UI_PKEY_ContextAvailable property of the Paint contextual tab Command and trigger the UpdatePropery callback function.
				g_pRibbonFramework->InvalidateUICommand(cmdContextualTabGroup_Paint, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_ContextAvailable);
			}
            break;

		case cmdHome_Commands_Run:
			{
				PROPVARIANT pVar;
				PropVariantInit(&pVar);
				BOOL RunEnabled = FALSE;

				g_pRibbonFramework->GetUICommandProperty(cmdHome_Commands_Run, UI_PKEY_Enabled, &pVar);

				HRESULT hr = PropVariantToBoolean(pVar, &RunEnabled);

				if (RunEnabled)
				{
					g_Variables.Running = !g_Variables.Running;

					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_ALLPROPERTIES, NULL);

					/*g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_BooleanValue);
					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_Label);
					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_LargeImage);
					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_SmallImage);*/
				}
			}
			break;

		case cmdHome_Commands_NextStep:
			{
				PROPVARIANT pVar;
				PropVariantInit(&pVar);
				BOOL NextStepEnabled = FALSE;

				g_pRibbonFramework->GetUICommandProperty(cmdHome_Commands_NextStep, UI_PKEY_Enabled, &pVar);

				HRESULT hr = PropVariantToBoolean(pVar, &NextStepEnabled);

				if (NextStepEnabled)
				{
					g_Variables.Running = true;
					g_Variables.NumberOfSteps = 1;
				}
			}
			break;

		case cmdHome_Commands_Clear:
			{
				PROPVARIANT pVar;
				PropVariantInit(&pVar);
				BOOL ClearEnabled = FALSE;

				g_pRibbonFramework->GetUICommandProperty(cmdHome_Commands_Clear, UI_PKEY_Enabled, &pVar);

				HRESULT hr = PropVariantToBoolean(pVar, &ClearEnabled);

				if (ClearEnabled)
				{
					if (g_Variables.Generation != 0U)
					{
						MessageBeep(MB_ICONQUESTION);

						if (IDNO == MessageBox(hWnd, L"Clear all existing grids?", L"Clear", MB_ICONQUESTION | MB_YESNO | MB_DEFBUTTON2 | MB_SETFOREGROUND | MB_TASKMODAL))
						{
							break;
						}
					}

					g_pTaskbarList->SetProgressState(hWnd, TBPF_INDETERMINATE);

					Concurrency::parallel_for(0, (int)g_Variables.Height, [&](UINT i)
					{
						for (UINT j = 0U; j < g_Variables.Width; j++)
						{
							g_Variables.EntropyCurrent[i][j] = 0U;
						}
					});

					g_Variables.EcosystemCurrent.clear();

					g_Variables.NewRun = true;
					g_Variables.Generation = 0;
					g_Variables.FreeEnergyInitial = 0;

					g_Variables.Redraw = true;
					g_Variables.Population = 0;
					g_Variables.ThermodynamicEntropy = 0;
					UpdateStatusBar();

					g_Variables.EntropyInitialised = false;
					g_Variables.LifeInitialised = false;
					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_NextStep, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Clear, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearEntropyGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearLifeGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
					g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearSplitButton, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);

					g_pTaskbarList->SetProgressState(hWnd, TBPF_NOPROGRESS);
				}
			}
			break;

		case cmdHome_Commands_ClearEntropyGrid:
			{
				if (g_Variables.Generation != 0)
				{
					MessageBeep(MB_ICONQUESTION);

					if (IDNO == MessageBox(hWnd, L"Clear entropy grid?", L"Clear", MB_ICONQUESTION | MB_YESNO | MB_DEFBUTTON2 | MB_SETFOREGROUND | MB_TASKMODAL))
					{
						break;
					}
				}

				g_pTaskbarList->SetProgressState(hWnd, TBPF_INDETERMINATE);

				Concurrency::parallel_for(0, (int)g_Variables.Height, [&](UINT i)
				{
					for (UINT j = 0; j < g_Variables.Width; j++)
					{
						g_Variables.EntropyCurrent[i][j] = 0;
					}
				});

				g_Variables.NewRun = true;
				g_Variables.FreeEnergyInitial = 0;

				g_Variables.Redraw = true;
				g_Variables.ThermodynamicEntropy = 0;
				UpdateStatusBar();

				g_Variables.EntropyInitialised = false;

				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_NextStep, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Clear, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearEntropyGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearLifeGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearSplitButton, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);

				g_pTaskbarList->SetProgressState(hWnd, TBPF_NOPROGRESS);
			}
			break;

		case cmdHome_Commands_ClearLifeGrid:
			{
				if (g_Variables.Generation != 0)
				{
					MessageBeep(MB_ICONQUESTION);

					if (IDNO == MessageBox(hWnd, L"Clear life grid?", L"Clear", MB_ICONQUESTION | MB_YESNO | MB_DEFBUTTON2 | MB_SETFOREGROUND | MB_TASKMODAL))
					{
						break;
					}
				}

				g_pTaskbarList->SetProgressState(hWnd, TBPF_INDETERMINATE);

				g_Variables.EcosystemCurrent.clear();

				g_Variables.NewRun = true;

				g_Variables.Redraw = true;
				g_Variables.Population = 0;
				UpdateStatusBar();

				g_Variables.LifeInitialised = false;

				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_NextStep, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Clear, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearEntropyGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearLifeGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearSplitButton, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);

				g_pTaskbarList->SetProgressState(hWnd, TBPF_NOPROGRESS);
			}
			break;

		case cmdHome_Ruleset_PeriodicBoundaryConditions:
			{
				g_Variables.PeriodicBoundaryConditions = !g_Variables.PeriodicBoundaryConditions;

				g_pRibbonFramework->InvalidateUICommand(cmdHome_Ruleset_PeriodicBoundaryConditions, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_BooleanValue);
			}
            break;

		case cmdHome_EntropyGrid_Initialise:
			{
				if (g_Variables.Generation != 0)
				{
					//MessageBeep(MB_ICONQUESTION);

					if (IDNO == MessageBox(hWnd, L"Overwrite the existing entropy grid?", L"Initialise", MB_ICONQUESTION | MB_YESNO | MB_DEFBUTTON2 | MB_SETFOREGROUND | MB_TASKMODAL))
					{
						break;
					}
				}

				/*if ((g_Variables.Initialisation == INITIALISATION_UNIFORM && g_Variables.UniformMin == 0 && g_Variables.UniformMax == 0) || (g_Variables.Initialisation == INITIALISATION_NORMAL && g_Variables.NormalMean == 0 && g_Variables.NormalVariance == 0))
				{
					break;
				}*/

				g_Variables.FreeEnergyInitial = 0;

				g_pTaskbarList->SetProgressState(hWnd, TBPF_INDETERMINATE);

				if (g_Variables.PeriodicBoundaryConditions)
				{
					if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
					{
						UINT random;

						for (UINT i = 0; i < g_Variables.Height; i++)
						{
							for (UINT j = 0; j < g_Variables.Width; j++)
							{
								random = RandomNumber::Uniform(g_Variables.RandomNumberFunction, g_Variables.UniformMin, g_Variables.UniformMax);

								g_Variables.EntropyCurrent[i][j] = random;
								g_Variables.FreeEnergyInitial = g_Variables.FreeEnergyInitial + random;
							}
						}
					}
					else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
					{
						double random;

						for (UINT i = 0; i < g_Variables.Height; i++)
						{
							for (UINT j = 0; j < g_Variables.Width; j++)
							{
								random = RandomNumber::Normal(g_Variables.RandomNumberFunction, g_Variables.NormalMean, g_Variables.NormalVariance);

								if (random < 0)
								{
									random = 0;
								}
								else if ((UINT)random > 255)
								{
									random = 255;
								}

								g_Variables.EntropyCurrent[i][j] = (UINT)random;
								g_Variables.FreeEnergyInitial = g_Variables.FreeEnergyInitial + (UINT)random;
							}
						}
					}
				}
				else
				{
					//Create borders
					for (UINT i = 0; i < g_Variables.Height; i++)
					{
						g_Variables.EntropyCurrent[i][0] = 0;
						g_Variables.EntropyCurrent[i][g_Variables.Width-1] = 0;
					}

					for (UINT j = 0; j < g_Variables.Width; j++)
					{
						g_Variables.EntropyCurrent[0][j] = 0;
						g_Variables.EntropyCurrent[g_Variables.Height-1][j] = 0;
					}

					if (g_Variables.Initialisation == INITIALISATION_UNIFORM)
					{
						UINT random;

						for (UINT i = 1; i < g_Variables.Height-1; i++)
						{
							for (UINT j = 1; j < g_Variables.Width-1; j++)
							{
								random = RandomNumber::Uniform(g_Variables.RandomNumberFunction, g_Variables.UniformMin, g_Variables.UniformMax);

								g_Variables.EntropyCurrent[i][j] = random;
								g_Variables.FreeEnergyInitial = g_Variables.FreeEnergyInitial + random;
							}
						}
					}
					else if (g_Variables.Initialisation == INITIALISATION_NORMAL)
					{
						double random;

						for (UINT i = 1; i < g_Variables.Height-1; i++)
						{
							for (UINT j = 1; j < g_Variables.Width-1; j++)
							{
								random = RandomNumber::Normal(g_Variables.RandomNumberFunction, g_Variables.NormalMean, g_Variables.NormalVariance);

								if (random < 0)
								{
									random = 0;
								}
								else if ((UINT)random > 255)
								{
									random = 255;
								}

								g_Variables.EntropyCurrent[i][j] = (UINT)random;
								g_Variables.FreeEnergyInitial = g_Variables.FreeEnergyInitial + (UINT)random;
							}
						}
					}
				}

				g_Variables.FreeEnergy = g_Variables.FreeEnergyInitial;

				g_Variables.Redraw = true;
				CalculateThermodynamicEntropy();
				UpdateStatusBar();

				g_Variables.EntropyInitialised = true;
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_NextStep, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Clear, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearEntropyGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearLifeGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearSplitButton, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);

				g_pTaskbarList->SetProgressState(hWnd, TBPF_NOPROGRESS);
			}
			break;

		case cmdHome_LifeGrid_Intelligence:
			{
				g_Variables.Intelligence = !g_Variables.Intelligence;

				g_pRibbonFramework->InvalidateUICommand(cmdHome_LifeGrid_Intelligence, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_BooleanValue);
			}
            break;

		case cmdHome_LifeGrid_Bias:
			{
				g_Variables.Bias = !g_Variables.Bias;

				g_pRibbonFramework->InvalidateUICommand(cmdHome_LifeGrid_Bias, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_BooleanValue);
			}
            break;

		case cmdHome_LifeGrid_Initialise:
			{
				if (g_Variables.Generation != 0)
				{
					//MessageBeep(MB_ICONQUESTION);

					if (IDNO == MessageBox(hWnd, L"Overwrite the existing life grid?", L"Initialise", MB_ICONQUESTION | MB_YESNO | MB_DEFBUTTON2 | MB_SETFOREGROUND | MB_TASKMODAL))
					{
						break;
					}
				}

				g_pTaskbarList->SetProgressState(hWnd, TBPF_INDETERMINATE);

				g_Variables.EcosystemCurrent.clear();

				g_Variables.EcosystemCurrent.resize(g_Variables.PopulationInitial);

				for (std::list <LIFEPROPERTIES>::iterator Iter = g_Variables.EcosystemCurrent.begin(); Iter != g_Variables.EcosystemCurrent.end(); Iter++)
				{
					if (g_Variables.PeriodicBoundaryConditions)
					{
						Iter->x = RandomNumber::Uniform(g_Variables.RandomNumberFunction, 0, g_Variables.Width);
						Iter->y = RandomNumber::Uniform(g_Variables.RandomNumberFunction, 0, g_Variables.Height);
					}
					else
					{
						Iter->x = RandomNumber::Uniform(g_Variables.RandomNumberFunction, 1, g_Variables.Width-1);
						Iter->y = RandomNumber::Uniform(g_Variables.RandomNumberFunction, 1, g_Variables.Height-1);
					}

					Iter->sex = (SEX)RandomNumber::Uniform(g_Variables.RandomNumberFunction, 0, 2);
					Iter->Lifespan = 100;
				};

				g_Variables.Population = g_Variables.PopulationInitial;

				g_Variables.Redraw = true;
				UpdateStatusBar();

				g_Variables.LifeInitialised = true;
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Run, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_NextStep, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_Clear, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearEntropyGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearLifeGrid, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);
				g_pRibbonFramework->InvalidateUICommand(cmdHome_Commands_ClearSplitButton, UI_INVALIDATIONS_STATE, &UI_PKEY_Enabled);

				g_pTaskbarList->SetProgressState(hWnd, TBPF_NOPROGRESS);
			}
			break;

		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}
		return 0;

	case WM_CREATE:
		{
			if (!InitializeRibbonFramework(hWnd))
			{
				return -1;
			}

			// Initialize common controls
			InitCommonControls();

			// Create status bar
			hStatusBar = CreateWindowEx(0, STATUSCLASSNAME, NULL, WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP, 0, 0, 0, 0, hWnd, (HMENU)IDC_STATUSBAR, GetModuleHandle(NULL), NULL);

			// Create status bar "compartments" at specified widths, last -1 means that it fills the rest of the window
			int StatusBarWidths[] = {g_metrics.ScaleX(128), g_metrics.ScaleX(256), g_metrics.ScaleX(384), g_metrics.ScaleX(512), -1};

			SendMessage(hStatusBar, SB_SETPARTS, (WPARAM)(sizeof(StatusBarWidths)/sizeof(int)), (LPARAM)StatusBarWidths);

			HANDLE handleIconCursorLocation = LoadImage(GetModuleHandle(NULL), MAKEINTRESOURCE(IDI_CURSORLOCATION), IMAGE_ICON, 16, 16, 0);

			SendMessage(hStatusBar, SB_SETICON, (WPARAM) 0, (LPARAM) handleIconCursorLocation);

			UpdateStatusBar();

			hRenderTarget = CreateWindowExW(
				0,
				L"RenderTarget",
				szTitle,
				WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
				CW_USEDEFAULT,
				0,
				CW_USEDEFAULT,
				0,
				hWnd,
				(HMENU)IDC_RENDERTARGET,
				hInst,
				NULL);

			if (!hRenderTarget)
			{
				ErrorDescription(HRESULT_FROM_WIN32(GetLastError()));
			}
			else
			{
				renderer.SetHwnd(hRenderTarget);
			}

			RegisterDropWindow(hWnd, &pDropTarget);
		}
		return 0;

	case WM_DESTROY:
		{
			SafeRelease(&g_pTaskbarList);
			DestroyRibbonFramework();
			UnregisterDropWindow(hWnd, pDropTarget);
			PostQuitMessage(0);
		}
		return 0;

	case WM_ERASEBKGND: //To prevent flicker
		{
		}
		return 1;

	/*case WM_GETMINMAXINFO:
		{
			//MessageBox(NULL, L"Access", L"Got here", MB_OK);

			MINMAXINFO * mmiStruct = (MINMAXINFO*)lParam;
			RECT rcDesktop, rcWind, rcClient, rcStatusBar;
			POINT ptDiff;
			POINT ptPoint;

			GetClientRect(hWnd, &rcClient);
			GetWindowRect(hWnd, &rcWind);
			ptDiff.x = (rcWind.right - rcWind.left) - rcClient.right;
			ptDiff.y = (rcWind.bottom - rcWind.top) - rcClient.bottom;

			GetClientRect(hStatusBar, &rcStatusBar);

			LONG desiredMinWidth = g_metrics.ScaleX(g_Variables.Width) + ptDiff.x;
			LONG desiredMinHeight = g_metrics.ScaleY(g_Variables.Height) + GetRibbonHeight() + (rcStatusBar.bottom - rcStatusBar.top) + ptDiff.y;

			SystemParametersInfo(SPI_GETWORKAREA, 0,  &rcDesktop, 0);

			if (desiredMinWidth > (rcDesktop.right - rcDesktop.left))
			{
				desiredMinWidth = (rcDesktop.right - rcDesktop.left;
			}

			if (desiredMinHeight > (rcDesktop.bottom - rcDesktop.top))
			{
				desiredMinHeight = (rcDesktop.bottom - rcDesktop.top)
			}

			ptPoint.x = desiredMinWidth;
			ptPoint.y = desiredMinHeight;

			mmiStruct->ptMinTrackSize = ptPoint;

			//ptPoint.x = GetSystemMetrics(SM_CXMAXIMIZED);	//Maximum width of the window.
			//ptPoint.y = GetSystemMetrics(SM_CYMAXIMIZED);	//Maximum height of the window.
			//mmiStruct->ptMaxTrackSize = ptPoint;
		}
		return 0;*/

	case WM_HELP:
		{
			
		}
		return true;

	case WM_DISPLAYCHANGE:
	case WM_PAINT:
		{
			hdc = BeginPaint(hWnd, &ps);

			EndPaint(hWnd, &ps);
		}
		return 0;

	/*case WM_SHOWWINDOW:
		{
			if ((BOOL)wParam == TRUE) //If wParam is TRUE, the window is being shown. If wParam is FALSE, the window is being hidden.
			{
				MessageBox(hWnd, L"Showing", L"WM_SHOWWINDOW", MB_OK);
				g_Variables.Redraw = true;
			}
		}
		return 0;*/

	case WM_SIZE:
		{
			UINT width = LOWORD(lParam);
            UINT height = HIWORD(lParam);

			// Auto-resize statusbar
			SendMessage(hStatusBar, WM_SIZE, 0, 0);

			RECT rc = {0};
			GetClientRect(hStatusBar, &rc);

			MoveWindow(hRenderTarget, 0, GetRibbonHeight(), width, (height - GetRibbonHeight() - (rc.bottom-rc.top)), true);
		}
		return 0;

	case WM_SIZING:
		{
			RECT* lprc = (LPRECT)lParam; // screen coordinates of drag rectangle

			RECT rcDesktop, rcWind, rcClient, rcStatusBar = {0};
			POINT ptDiff = {0};

			GetClientRect(hWnd, &rcClient);
			GetWindowRect(hWnd, &rcWind);
			ptDiff.x = (rcWind.right - rcWind.left) - rcClient.right;
			ptDiff.y = (rcWind.bottom - rcWind.top) - rcClient.bottom;

			GetClientRect(hStatusBar, &rcStatusBar);

			LONG desiredMinWidth = g_metrics.ScaleX(g_Variables.Width) + ptDiff.x;
			LONG desiredMinHeight = g_metrics.ScaleY(g_Variables.Height) + GetRibbonHeight() + (rcStatusBar.bottom - rcStatusBar.top) + ptDiff.y;

			SystemParametersInfo(SPI_GETWORKAREA, 0,  &rcDesktop, 0);

			if (desiredMinWidth > (rcDesktop.right - rcDesktop.left))
			{
				desiredMinWidth = rcDesktop.right - rcDesktop.left;
			}

			if (desiredMinHeight > (rcDesktop.bottom - rcDesktop.top))
			{
				desiredMinHeight = rcDesktop.bottom - rcDesktop.top;
			}

			switch (wParam)
			{
			case WMSZ_TOP:
				{
					if ((lprc->bottom - lprc->top) < desiredMinHeight)
					{
						lprc->top = lprc->bottom - desiredMinHeight;
					}
				}
				break;
			case WMSZ_TOPRIGHT:
				{
					if ((lprc->right - lprc->left) < desiredMinWidth)
					{
						lprc->right = lprc->left + desiredMinWidth;
				
					}

					if ((lprc->bottom - lprc->top) < desiredMinHeight)
					{
						lprc->top = lprc->bottom - desiredMinHeight;
					}
				}
				break;
			case WMSZ_RIGHT:
				{
					if ((lprc->right - lprc->left) < desiredMinWidth)
					{
						lprc->right = lprc->left + desiredMinWidth;
				
					}
				}
				break;
			case WMSZ_BOTTOMRIGHT:
				{
					if ((lprc->right - lprc->left) < desiredMinWidth)
					{
						lprc->right = lprc->left + desiredMinWidth;
				
					}

					if ((lprc->bottom - lprc->top) < desiredMinHeight)
					{
						lprc->bottom = lprc->top + desiredMinHeight;
					}
				}
				break;
			case WMSZ_BOTTOM:
				{
					if ((lprc->bottom - lprc->top) < desiredMinHeight)
					{
						lprc->bottom = lprc->top + desiredMinHeight;
					}
				}
				break;
			case WMSZ_BOTTOMLEFT:
				{
					if ((lprc->bottom - lprc->top) < desiredMinHeight)
					{
						lprc->bottom = lprc->top + desiredMinHeight;
					}

					if ((lprc->right - lprc->left) < desiredMinWidth)
					{
						lprc->left = lprc->right - desiredMinWidth;
				
					}
				}
				break;
			case WMSZ_LEFT:
				{
					if ((lprc->right - lprc->left) < desiredMinWidth)
					{
						lprc->left = lprc->right - desiredMinWidth;
				
					}
				}
				break;
			case WMSZ_TOPLEFT:
				{
					if ((lprc->bottom - lprc->top) < desiredMinHeight)
					{
						lprc->top = lprc->bottom - desiredMinHeight;
					}

					if ((lprc->right - lprc->left) < desiredMinWidth)
					{
						lprc->left = lprc->right - desiredMinWidth;
				
					}
				}
				break;
			}
		}
		return true;

	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}