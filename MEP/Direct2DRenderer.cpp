#include "stdafx.h"
#include "Direct2DRenderer.h"

extern VARIABLES g_Variables;

inline void Direct2DRenderer::CreateBitmapData()
{
	Concurrency::parallel_for(0, (int)g_Variables.Height, [&](UINT i)
	{
		for (UINT j = 0; j < g_Variables.Width; j++)
		{
			UINT index = (j * (g_Variables.BytesPerPixel)) + (i * g_Variables.Stride);
			
			g_Variables.BitmapDataEntropy[index]		= (byte)(g_Variables.EntropyCurrent[i][j]); //Blue
			g_Variables.BitmapDataEntropy[(index + 1)]	= (byte)(g_Variables.EntropyCurrent[i][j]); //Green
			g_Variables.BitmapDataEntropy[(index + 2)]	= (byte)(g_Variables.EntropyCurrent[i][j]); //Red
			g_Variables.BitmapDataEntropy[(index + 3)]	= (byte)(255); //Alpha

			g_Variables.BitmapDataLife[index]		= (byte)(0); //Blue
			g_Variables.BitmapDataLife[(index + 1)]	= (byte)(0); //Green
			g_Variables.BitmapDataLife[(index + 2)]	= (byte)(0); //Red
			g_Variables.BitmapDataLife[(index + 3)]	= (byte)(0); //Alpha
		}
	});

	/*for (UINT i = 0; i < g_Variables.Height; i++)
	{
		for (UINT j = 0; j < g_Variables.Width; j++)
		{
			UINT index = (j * (g_Variables.BytesPerPixel)) + (i * g_Variables.Stride);
			
			g_Variables.BitmapDataEntropy[index]		= (byte)(g_Variables.EntropyCurrent[i][j]); //Blue
			g_Variables.BitmapDataEntropy[(index + 1)]	= (byte)(g_Variables.EntropyCurrent[i][j]); //Green
			g_Variables.BitmapDataEntropy[(index + 2)]	= (byte)(g_Variables.EntropyCurrent[i][j]); //Red
			g_Variables.BitmapDataEntropy[(index + 3)]	= (byte)(255); //Alpha

			g_Variables.BitmapDataLife[index]		= (byte)(0); //Blue
			g_Variables.BitmapDataLife[(index + 1)]	= (byte)(0); //Green
			g_Variables.BitmapDataLife[(index + 2)]	= (byte)(0); //Red
			g_Variables.BitmapDataLife[(index + 3)]	= (byte)(0); //Alpha
		}
	};*/

	 for (std::list <LIFEPROPERTIES>::iterator Iter = g_Variables.EcosystemCurrent.begin(); Iter != g_Variables.EcosystemCurrent.end(); Iter++)
	 {
		 UINT index = (Iter->x * (g_Variables.BytesPerPixel)) + (Iter->y * g_Variables.Stride);

		 if (Iter->sex == MALE)
		 {
			 g_Variables.BitmapDataLife[(index)] = (byte)(255); //Blue
		 }
		 else
		 {
			  g_Variables.BitmapDataLife[(index + 2)] = (byte)(255); //Red
		 }
		 
		 g_Variables.BitmapDataLife[(index + 3)] = (byte)(255); //Alpha
	 }
}

//
// Initialize members.
//
Direct2DRenderer::Direct2DRenderer() :
    m_hwnd(NULL),
    m_pD2DFactory(NULL),
    m_pWICFactory(NULL),
    m_pRenderTarget(NULL),
	m_pBitmapEntropy(NULL),
	m_pBitmapLife(NULL)
{
}

//
// Release resources.
//
Direct2DRenderer::~Direct2DRenderer()
{
	//SafeRelease(&m_pWICFactory); Already destroyed by CoUninitialize() call
    SafeRelease(&m_pD2DFactory);
    SafeRelease(&m_pRenderTarget);
    SafeRelease(&m_pBitmapEntropy);
	SafeRelease(&m_pBitmapLife);
}

//
// Create resources which are not bound
// to any device. Their lifetime effectively extends for the
// duration of the app. These resources include the Direct2D,
// DirectWrite, and WIC factories; and a DirectWrite Text Format object
// (used for identifying particular font characteristics) and
// a Direct2D geometry.
//
HRESULT Direct2DRenderer::CreateDeviceIndependentResources()
{
    HRESULT hr;

    // Create a Direct2D factory.
    hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, &m_pD2DFactory);
    if (SUCCEEDED(hr))
    {
        // Create WIC factory.
        hr = CoCreateInstance(
            CLSID_WICImagingFactory,
            NULL,
            CLSCTX_INPROC_SERVER,
            IID_IWICImagingFactory,
            reinterpret_cast<void **>(&m_pWICFactory)
            );
    }

    return hr;
}

//
//  This method creates resources which are bound to a particular
//  Direct3D device. It's all centralized here, in case the resources
//  need to be recreated in case of Direct3D device loss (eg. display
//  change, remoting, removal of video card, etc).
//
HRESULT Direct2DRenderer::CreateDeviceResources()
{
    HRESULT hr = S_OK;

    if (!m_pRenderTarget)
    {
		//MessageBox(NULL, L"CreateDeviceResources", L"Info", MB_ICONINFORMATION | MB_OK);

        RECT rc;
        GetClientRect(m_hwnd, &rc);

        D2D1_SIZE_U size = D2D1::SizeU(
            rc.right - rc.left,
            rc.bottom - rc.top
            );

        // Create a Direct2D render target.
        hr = m_pD2DFactory->CreateHwndRenderTarget(
            D2D1::RenderTargetProperties(),
            D2D1::HwndRenderTargetProperties(m_hwnd, size),
            &m_pRenderTarget
            );

		m_pD2DFactory->GetDesktopDpi(&m_dpiX, &m_dpiY);

		hr = m_pRenderTarget->CreateBitmap(
			D2D1::SizeU(g_Variables.Width, g_Variables.Height),
			NULL,
			NULL,
			D2D1::BitmapProperties(D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED), m_dpiX, m_dpiY),
			&m_pBitmapEntropy
			);

		hr = m_pRenderTarget->CreateBitmap(
			D2D1::SizeU(g_Variables.Width, g_Variables.Height),
			NULL,
			NULL,
			D2D1::BitmapProperties(D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED), m_dpiX, m_dpiY),
			&m_pBitmapLife
			);
    }

    return hr;
}

//
//  Discard device-specific resources which need to be recreated
//  when a Direct3D device is lost
//
void Direct2DRenderer::DiscardDeviceResources()
{
    SafeRelease(&m_pRenderTarget);
    SafeRelease(&m_pBitmapEntropy);
	SafeRelease(&m_pBitmapLife);
}

//
//  Called whenever the application needs to display the client
//  window.
//
//  Note that this function will not render anything if the window
//  is occluded (e.g. when the screen is locked).
//  Also, this function will automatically discard device-specific
//  resources if the Direct3D device disappears during function
//  invocation, and will recreate the resources the next time it's
//  invoked.
//

HRESULT Direct2DRenderer::OnRender()
{
    HRESULT hr = CreateDeviceResources();

    if (SUCCEEDED(hr) && !(m_pRenderTarget->CheckWindowState() & D2D1_WINDOW_STATE_OCCLUDED))
    {
        // Retrieve the size of the render target.
        //D2D1_SIZE_F renderTargetSize = m_pRenderTarget->GetSize();

		if (g_Variables.Redraw)
		{
			CreateBitmapData();
			
			hr = m_pBitmapEntropy->CopyFromMemory(NULL, g_Variables.BitmapDataEntropy, g_Variables.Stride);
			if (FAILED(hr))
			{
				return hr;
				//MessageBeep(MB_ICONERROR);
				//MessageBox(NULL, L"CopyFromMemory failed (Entropy)", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
			}

			hr = m_pBitmapLife->CopyFromMemory(NULL, g_Variables.BitmapDataLife, g_Variables.Stride);
			if (FAILED(hr))
			{
				return hr;
				//MessageBeep(MB_ICONERROR);
				//MessageBox(NULL, L"CopyFromMemory failed (Life)", L"Error", MB_ICONERROR | MB_OK | MB_TASKMODAL);
			}

			g_Variables.Redraw = false;
		}

        m_pRenderTarget->BeginDraw();

        //m_pRenderTarget->SetTransform(D2D1::Matrix3x2F::Identity());
		//m_pRenderTarget->SetTransform(D2D1::Matrix3x2F::Translation(50, 50));

        m_pRenderTarget->Clear(D2D1::ColorF(D2D1::ColorF::Gainsboro));

		D2D1_SIZE_F size = m_pBitmapEntropy->GetSize();

        m_pRenderTarget->DrawBitmap(
            m_pBitmapEntropy,
			D2D1::RectF((FLOAT)g_Variables.TranslateVector.x, (FLOAT)g_Variables.TranslateVector.y, (m_dpiX/96.0F)*size.width + g_Variables.TranslateVector.x, (m_dpiY/96.0F)*size.height + g_Variables.TranslateVector.y)
            );

		m_pRenderTarget->DrawBitmap(
            m_pBitmapLife,
			D2D1::RectF((FLOAT)g_Variables.TranslateVector.x, (FLOAT)g_Variables.TranslateVector.y, (m_dpiX/96.0F)*size.width + g_Variables.TranslateVector.x, (m_dpiY/96.0F)*size.height + g_Variables.TranslateVector.y)
            );

        hr = m_pRenderTarget->EndDraw();

        if (hr == D2DERR_RECREATE_TARGET)
        {
            hr = S_OK;
            DiscardDeviceResources();
        }
    }

    return hr;
}

HRESULT Direct2DRenderer::OnClear()
{
    HRESULT hr;

    hr = CreateDeviceResources();

    if (SUCCEEDED(hr) && !(m_pRenderTarget->CheckWindowState() & D2D1_WINDOW_STATE_OCCLUDED))
    {
        m_pRenderTarget->BeginDraw();

        m_pRenderTarget->Clear(D2D1::ColorF(D2D1::ColorF::Gainsboro));

        hr = m_pRenderTarget->EndDraw();

        if (hr == D2DERR_RECREATE_TARGET)
        {
            hr = S_OK;
            DiscardDeviceResources();
        }
    }

    return hr;
}

//
//  If the application receives a WM_SIZE message, this method
//  resize the render target appropriately.
//
void Direct2DRenderer::OnResize(UINT width, UINT height)
{
    if (m_pRenderTarget)
    {
        // Note: This method can fail, but it's okay to ignore the
        // error here -- it will be repeated on the next call to
        // EndDraw.
		m_pRenderTarget->Resize(D2D1::SizeU(width, height));
    }
}

HRESULT Direct2DRenderer::CreateFileFromLayer(PCWSTR uri, LAYER layer, GUID guidContainerFormat)
{
	return CreateFileFromLayer(m_pWICFactory, uri, g_Variables.Width, g_Variables.Height, layer, guidContainerFormat);
};

HRESULT Direct2DRenderer::CreateFileFromLayer(
    IWICImagingFactory *pIWICFactory,
    PCWSTR uri,
    UINT destinationWidth,
    UINT destinationHeight,
	LAYER layer,
	GUID guidContainerFormat
    )
{
    HRESULT hr = S_OK;

	//Variables used for encoding.
IWICStream *piFileStream = NULL;
IWICBitmapEncoder *piEncoder = NULL;

WICPixelFormatGUID pixelFormat = GUID_WICPixelFormat32bppBGRA;

// Create a file stream.
if (SUCCEEDED(hr))
{
    hr = pIWICFactory->CreateStream(&piFileStream);
}

// Initialize our new file stream.
if (SUCCEEDED(hr))
{
    hr = piFileStream->InitializeFromFilename(uri, GENERIC_WRITE);
}

// Create the encoder.
if (SUCCEEDED(hr))
{
    hr = pIWICFactory->CreateEncoder(guidContainerFormat, NULL, &piEncoder);
}
// Initialize the encoder
if (SUCCEEDED(hr))
{
    hr = piEncoder->Initialize(piFileStream, WICBitmapEncoderNoCache);
}

IWICBitmapFrameEncode *piFrameEncode = NULL;

hr = piEncoder->CreateNewFrame(&piFrameEncode, NULL);

if (SUCCEEDED(hr))
    {
        hr = piFrameEncode->Initialize(NULL);
    }
if (SUCCEEDED(hr))
    {
        hr = piFrameEncode->SetSize(destinationWidth, destinationHeight);
    }
if (SUCCEEDED(hr))
    {
        hr = piFrameEncode->SetResolution(96.0F, 96.0F);
    }
if (SUCCEEDED(hr))
    {
        hr = piFrameEncode->SetPixelFormat(&pixelFormat);
    }
if (SUCCEEDED(hr))
    {
		if (layer == LAYER_ENTROPY)
		{
			hr = piFrameEncode->WritePixels(g_Variables.Height, g_Variables.Stride, g_Variables.Stride*g_Variables.Height, g_Variables.BitmapDataEntropy);

        /*hr = piFrameEncode->WriteSource(
            m_pWICBitmapEntropy,
            NULL);*/
		}
		else if (layer == LAYER_LIFE)
		{
			hr = piFrameEncode->WritePixels(g_Variables.Height, g_Variables.Stride, g_Variables.Stride*g_Variables.Height, g_Variables.BitmapDataLife);
			/*hr = piFrameEncode->WriteSource(
            m_pWICBitmapLife,
            NULL);*/
		}
    }

    //Commit the frame.
    if (SUCCEEDED(hr))
    {
        hr = piFrameEncode->Commit();
    }
if (SUCCEEDED(hr))
{
    piEncoder->Commit();
}

if (SUCCEEDED(hr))
{
    piFileStream->Commit(STGC_DEFAULT);
}


SafeRelease(&piFileStream);
SafeRelease(&piEncoder);
SafeRelease(&piFrameEncode);

    return hr;
}

HRESULT Direct2DRenderer::CreateLayerFromFile(PCWSTR uri, LAYER layer)
{
	return CreateLayerFromFile(m_pWICFactory, uri, g_Variables.Width, g_Variables.Height, layer);
};

//
// Creates a WICbitmap from the specified
// file name.
//
HRESULT Direct2DRenderer::CreateLayerFromFile(
    IWICImagingFactory *pIWICFactory,
    PCWSTR uri,
    UINT destinationWidth,
    UINT destinationHeight,
	LAYER layer
    )
{
    IWICBitmapDecoder *pDecoder = nullptr;
    IWICBitmapFrameDecode *pSource = nullptr;
    IWICStream *pStream = nullptr;
    IWICFormatConverter *pConverter = nullptr;
    IWICBitmapScaler *pScaler = nullptr;

    HRESULT hr = pIWICFactory->CreateDecoderFromFilename(
        uri,
        NULL,
        GENERIC_READ,
        WICDecodeMetadataCacheOnDemand,
        &pDecoder
        );

    if (SUCCEEDED(hr))
    {
        // Create the initial frame.
        hr = pDecoder->GetFrame(0, &pSource);
    }

    if (SUCCEEDED(hr))
    {
        // Convert the image format to 32bppPBGRA
        // (DXGI_FORMAT_B8G8R8A8_UNORM + D2D1_ALPHA_MODE_PREMULTIPLIED).
        hr = pIWICFactory->CreateFormatConverter(&pConverter);
    }

    if (SUCCEEDED(hr))
    {
        // If a new width or height was specified, create an
        // IWICBitmapScaler and use it to resize the image.
        if (destinationWidth != 0 || destinationHeight != 0)
        {
            UINT originalWidth, originalHeight;

            hr = pSource->GetSize(&originalWidth, &originalHeight);

            if (SUCCEEDED(hr))
            {
                if (destinationWidth == 0)
                {
                    FLOAT scalar = static_cast<FLOAT>(destinationHeight) / static_cast<FLOAT>(originalHeight);
                    destinationWidth = static_cast<UINT>(scalar * static_cast<FLOAT>(originalWidth));
                }
                else if (destinationHeight == 0)
                {
                    FLOAT scalar = static_cast<FLOAT>(destinationWidth) / static_cast<FLOAT>(originalWidth);
                    destinationHeight = static_cast<UINT>(scalar * static_cast<FLOAT>(originalHeight));
                }

                hr = pIWICFactory->CreateBitmapScaler(&pScaler);

                if (SUCCEEDED(hr))
                {
                    hr = pScaler->Initialize(
                            pSource,
                            destinationWidth,
                            destinationHeight,
                            WICBitmapInterpolationModeCubic
                            );
                }

                if (SUCCEEDED(hr))
                {
                    hr = pConverter->Initialize(
                        pScaler,
                        GUID_WICPixelFormat32bppPBGRA,
                        WICBitmapDitherTypeNone,
                        NULL,
                        0.f,
                        WICBitmapPaletteTypeMedianCut
                        );
                }
            }
        }
        else // Don't scale the image.
        {
            hr = pConverter->Initialize(
                pSource,
                GUID_WICPixelFormat32bppPBGRA,
                WICBitmapDitherTypeNone,
                NULL,
                0.f,
                WICBitmapPaletteTypeMedianCut
                );
        }
    }

    if (SUCCEEDED(hr))
    {
		
		if (layer == LAYER_ENTROPY)
		{
			hr = pConverter->CopyPixels(NULL, g_Variables.Stride, g_Variables.Stride*g_Variables.Height, g_Variables.BitmapDataEntropy);
		}
		else if (layer == LAYER_LIFE)
		{
			hr = pConverter->CopyPixels(NULL, g_Variables.Stride, g_Variables.Stride*g_Variables.Height, g_Variables.BitmapDataLife);
		}
    }

    SafeRelease(&pDecoder);
    SafeRelease(&pSource);
    SafeRelease(&pStream);
    SafeRelease(&pConverter);
    SafeRelease(&pScaler);

    return hr;
}

void Direct2DRenderer::SetHwnd(HWND hWnd)
{
	m_hwnd = hWnd;
}