#pragma once

#include <d2d1.h>
#include <d2d1helper.h>
#include <wincodec.h>

template<class Interface>
inline void SafeRelease(Interface **ppInterfaceToRelease)
{
    if (*ppInterfaceToRelease != NULL)
    {
        (*ppInterfaceToRelease)->Release();

        (*ppInterfaceToRelease) = NULL;
    }
}

class Direct2DRenderer
{
public:
    Direct2DRenderer();
    ~Direct2DRenderer();

	HRESULT CreateDeviceIndependentResources();

    HRESULT OnRender();
	HRESULT OnClear();

    void OnResize(UINT width, UINT height);

	HRESULT CreateLayerFromFile(PCWSTR uri, LAYER layer);
	HRESULT CreateFileFromLayer(PCWSTR uri, LAYER layer, GUID guidContainerFormat);

	void SetHwnd(HWND hWnd);

private:
	void CreateBitmapData();

    HRESULT CreateDeviceResources();
    void DiscardDeviceResources();

	HRESULT CreateFileFromLayer(
    IWICImagingFactory *pIWICFactory,
    PCWSTR uri,
    UINT destinationWidth,
    UINT destinationHeight,
    LAYER layer,
	GUID guidContainerFormat
    );

	HRESULT CreateLayerFromFile(
    IWICImagingFactory *pIWICFactory,
    PCWSTR uri,
    UINT destinationWidth,
    UINT destinationHeight,
    LAYER layer
    );

	HWND m_hwnd;
    ID2D1Factory *m_pD2DFactory;
    IWICImagingFactory *m_pWICFactory;
    ID2D1HwndRenderTarget *m_pRenderTarget;
	ID2D1Bitmap *m_pBitmapEntropy;
	ID2D1Bitmap *m_pBitmapLife;
	FLOAT m_dpiX, m_dpiY;
};