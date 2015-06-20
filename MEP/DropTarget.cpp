#include "stdafx.h"
#include "DropTarget.h"
#include "Direct2DRenderer.h"

extern void CreateLayerFromFile(PCWSTR FileName, LAYER layer);

//
//	Constructor for the CDropTarget class
//
CDropTarget::CDropTarget(HWND hWnd)
{
	m_lRefCount  = 1;
	m_hWnd       = hWnd;
	m_fAllowDrop = false;
}

//
//	Destructor for the CDropTarget class
//
CDropTarget::~CDropTarget()
{
	
}

//
//	IUnknown::QueryInterface
//
HRESULT __stdcall CDropTarget::QueryInterface (REFIID iid, void **ppvObject)
{
	if (iid == IID_IDropTarget || iid == IID_IUnknown)
	{
		AddRef();
		*ppvObject = this;
		return S_OK;
	}
	else
	{
		*ppvObject = 0;
		return E_NOINTERFACE;
	}
}

//
//	IUnknown::AddRef
//
ULONG __stdcall CDropTarget::AddRef(void)
{
	return InterlockedIncrement(&m_lRefCount);
}	

//
//	IUnknown::Release
//
ULONG __stdcall CDropTarget::Release(void)
{
	LONG count = InterlockedDecrement(&m_lRefCount);
		
	if (count == 0)
	{
		delete this;
		return 0;
	}
	else
	{
		return count;
	}
}

//
//	QueryDataObject private helper routine
//
bool CDropTarget::QueryDataObject(IDataObject *pDataObject)
{
	FORMATETC fmtetc = { CF_HDROP, 0, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };

	// does the data object support CF_TEXT using a HGLOBAL?
	if (pDataObject->QueryGetData(&fmtetc) == S_OK)
	{
		STGMEDIUM stgmed;

		if (pDataObject->GetData(&fmtetc, &stgmed) == S_OK)
		{
			TCHAR buffer[MAX_PATH];

			HDROP data = (HDROP)stgmed.hGlobal;

			DragQueryFile(data, 0, buffer, MAX_PATH);

			TCHAR *suffix = _tcsrchr(buffer, '.');

			if (suffix)
			{
			  suffix++;
			  if (_tcsicmp(suffix, TEXT("bmp")) == 0
				  || _tcsicmp(suffix, TEXT("dib")) == 0
				  || _tcsicmp(suffix, TEXT("gif")) == 0
				  || _tcsicmp(suffix, TEXT("jpg")) == 0
				  || _tcsicmp(suffix, TEXT("jpeg")) == 0
				  || _tcsicmp(suffix, TEXT("jpe")) == 0
				  || _tcsicmp(suffix, TEXT("jif")) == 0
				  || _tcsicmp(suffix, TEXT("jfif")) == 0
				  || _tcsicmp(suffix, TEXT("jfi")) == 0
				  || _tcsicmp(suffix, TEXT("png")) == 0
				  || _tcsicmp(suffix, TEXT("tif")) == 0
				  || _tcsicmp(suffix, TEXT("tiff")) == 0
				  )
			  {
				return true;
			  }
			}
		}
	}

	return false;
}

//
//	DropEffect private helper routine
//
DWORD CDropTarget::DropEffect(DWORD grfKeyState, POINTL /*pt*/, DWORD dwAllowed)
{
	DWORD dwEffect = 0;

	// 1. check "pt" -> do we allow a drop at the specified coordinates?
	
	// 2. work out that the drop-effect should be based on grfKeyState
	if (grfKeyState & MK_CONTROL)
	{
		dwEffect = dwAllowed & DROPEFFECT_COPY;
	}
	else if (grfKeyState & MK_SHIFT)
	{
		dwEffect = dwAllowed & DROPEFFECT_MOVE;
	}
	
	// 3. no key-modifiers were specified (or drop effect not allowed), so
	//    base the effect on those allowed by the dropsource
	if (dwEffect == 0)
	{
		if (dwAllowed & DROPEFFECT_COPY) dwEffect = DROPEFFECT_COPY;
		if (dwAllowed & DROPEFFECT_MOVE) dwEffect = DROPEFFECT_MOVE;
	}
	
	return dwEffect;
}


//
//	IDropTarget::DragEnter
//
//
//
HRESULT __stdcall CDropTarget::DragEnter(IDataObject *pDataObject, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	// does the dataobject contain data we want?
	m_fAllowDrop = QueryDataObject(pDataObject);
	
	if(m_fAllowDrop)
	{
		// get the dropeffect based on keyboard state
		*pdwEffect = DropEffect(grfKeyState, pt, *pdwEffect);

		SetFocus(m_hWnd);
	}
	else
	{
		*pdwEffect = DROPEFFECT_NONE;
	}

	return S_OK;
}

//
//	IDropTarget::DragOver
//
//
//
HRESULT __stdcall CDropTarget::DragOver(DWORD grfKeyState, POINTL pt, DWORD * pdwEffect)
{
	if(m_fAllowDrop)
	{
		*pdwEffect = DropEffect(grfKeyState, pt, *pdwEffect);
	}
	else
	{
		*pdwEffect = DROPEFFECT_NONE;
	}

	return S_OK;
}

//
//	IDropTarget::DragLeave
//
HRESULT __stdcall CDropTarget::DragLeave(void)
{
	return S_OK;
}

//
//	IDropTarget::Drop
//
//
HRESULT __stdcall CDropTarget::Drop(IDataObject * pDataObject, DWORD grfKeyState, POINTL pt, DWORD * pdwEffect)
{
	if (m_fAllowDrop)
	{
		DropData(m_hWnd, pDataObject);

		*pdwEffect = DropEffect(grfKeyState, pt, *pdwEffect);
	}
	else
	{
		*pdwEffect = DROPEFFECT_NONE;
	}

	return S_OK;
}

void DropData(HWND hWnd, IDataObject *pDataObject)
{
	// construct a FORMATETC object
	FORMATETC fmtetc = { CF_HDROP, 0, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	STGMEDIUM stgmed;

	// See if the dataobject contains any TEXT stored as a HGLOBAL
	if (pDataObject->QueryGetData(&fmtetc) == S_OK)
	{
		// The data is there, so allocate it
		if (pDataObject->GetData(&fmtetc, &stgmed) == S_OK)
		{
			// we asked for the data as a HGLOBAL, so access it appropriately
			HDROP data = (HDROP)stgmed.hGlobal;

			TCHAR buffer[MAX_PATH];

			DragQueryFile(data, 0, buffer, MAX_PATH);

			MessageBeep(MB_ICONQUESTION);
			int MessageBoxDecision = MessageBox(hWnd, TEXT("Paste to entropy grid?\nClicking \"No\" will paste to the life grid."), TEXT("Drag and drop"), MB_DEFBUTTON3 | MB_ICONQUESTION | MB_SETFOREGROUND | MB_TASKMODAL | MB_YESNOCANCEL);

			if (MessageBoxDecision == IDYES)
			{
				CreateLayerFromFile(buffer, LAYER_ENTROPY);

				//wcscat_s(buffer, MAX_PATH,   L" - Entropy");
				//SetWindowText(hWnd, buffer);
			}
			else if (MessageBoxDecision == IDNO)
			{
				CreateLayerFromFile(buffer, LAYER_LIFE);
				//wcscat_s(buffer, MAX_PATH,   L" - Life");
				//SetWindowText(hWnd, buffer);
			}

			// release the data using the COM API
			ReleaseStgMedium(&stgmed);
		}
	}
}

void RegisterDropWindow(HWND hWnd, IDropTarget **ppDropTarget)
{
	CDropTarget *pDropTarget = new CDropTarget(hWnd);

	// acquire a strong lock
	CoLockObjectExternal(pDropTarget, true, false);

	// tell OLE that the window is a drop target
	RegisterDragDrop(hWnd, pDropTarget);

	*ppDropTarget = pDropTarget;
}

void UnregisterDropWindow(HWND hWnd, IDropTarget *pDropTarget)
{
	// remove drag+drop
	RevokeDragDrop(hWnd);

	// remove the strong lock
	CoLockObjectExternal(pDropTarget, false, true);

	// release our own reference
	pDropTarget->Release();
}