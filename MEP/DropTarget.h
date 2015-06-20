
void DropData(HWND hWnd, IDataObject *pDataObject);
void RegisterDropWindow(HWND hWnd, IDropTarget **ppDropTarget);
void UnregisterDropWindow(HWND hWnd, IDropTarget *pDropTarget);

//
//	This is our definition of a class which implements
//  the IDropTarget interface
//
class CDropTarget : public IDropTarget
{
public:
	// IUnknown implementation
	HRESULT __stdcall QueryInterface (REFIID iid, void ** ppvObject);
	ULONG	__stdcall AddRef (void);
	ULONG	__stdcall Release (void);

	// IDropTarget implementation
	HRESULT __stdcall DragEnter (IDataObject * pDataObject, DWORD grfKeyState, POINTL pt, DWORD * pdwEffect);
	HRESULT __stdcall DragOver (DWORD grfKeyState, POINTL pt, DWORD * pdwEffect);
	HRESULT __stdcall DragLeave (void);
	HRESULT __stdcall Drop (IDataObject * pDataObject, DWORD grfKeyState, POINTL pt, DWORD * pdwEffect);

	// Constructor
	CDropTarget(HWND hWnd);
	~CDropTarget();

private:

	// internal helper function
	DWORD DropEffect(DWORD grfKeyState, POINTL pt, DWORD dwAllowed);
	bool  QueryDataObject(IDataObject *pDataObject);


	// Private member variables
	LONG	m_lRefCount;
	HWND	m_hWnd;
	bool    m_fAllowDrop;

	IDataObject *m_pDataObject;

};