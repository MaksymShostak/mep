#pragma once

#include <time.h> //time(0)
#include <algorithm> //std::for_each
#include <fstream> //std::basic_ofstream
#include <propvarutil.h> //PropVariantToBoolean
#include <shobjidl.h> //ITaskbarList3
//#include <math.h>

#include "resource.h"
#include "ribbonres.h"

ATOM				MyRegisterClass(HINSTANCE hInstance);
ATOM				MyRegisterRenderTargetClass(HINSTANCE hInstance);
bool				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK	RenderTargetWndProc(HWND, UINT, WPARAM, LPARAM);
void				ClientResize(HWND &hWnd, UINT nWidth, UINT nHeight);
void				RunSimulation();
void				CalculatePopulation();
void				CalculateThermodynamicEntropy();
void				UpdateStatusBar();
void				SetDlgItemUrl(HWND hdlg, int id, const char *url);