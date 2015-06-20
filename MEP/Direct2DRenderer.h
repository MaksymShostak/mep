#pragma once

#include <d2d1.h>
#include <d2d1helper.h>
#include <wincodec.h>

class Direct2DRenderer
{
public:
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

	HWND m_hwnd = nullptr;
	Microsoft::WRL::ComPtr<ID2D1Factory> m_pD2DFactory;
	Microsoft::WRL::ComPtr<IWICImagingFactory> m_pWICFactory;
	Microsoft::WRL::ComPtr<ID2D1HwndRenderTarget> m_pRenderTarget;
	Microsoft::WRL::ComPtr<ID2D1Bitmap> m_pBitmapEntropy;
	Microsoft::WRL::ComPtr<ID2D1Bitmap> m_pBitmapLife;
	FLOAT m_dpiX = 96.0F;
	FLOAT m_dpiY = 96.0F;
};