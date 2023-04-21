#pragma once

//#include "iunknownimpl.h"

//#include "smartrenameinterfaces.h"
//#include <shldisp.h>
//#include <atomic>

class CSmartRenameListView
{
public:
	CSmartRenameListView(void) noexcept = default;
	~CSmartRenameListView(void) noexcept = default;

	void Init(_In_ HWND hwndLV);
	void ToggleAll(_In_ ISmartRenameManager* psrm, _In_ bool selected);
	void ToggleItem(_In_ ISmartRenameManager* psrm, _In_ int item);
	void UpdateItemCheckState(_In_ ISmartRenameManager* psrm, _In_ int iItem);
	void RedrawItems(_In_ uint32_t first, _In_ uint32_t last);
	void SetItemCount(_In_ uint32_t itemCount);
	void OnKeyDown(_In_ ISmartRenameManager* psrm, _In_ NMLVKEYDOWN* lvKeyDown);
	void OnClickList(_In_ ISmartRenameManager* psrm, NMLISTVIEW* pnmListView);
	void GetDisplayInfo(_In_ ISmartRenameManager* psrm, _Inout_ NMLVDISPINFOW* plvdi);
	void OnSize();
	HWND GetHWND()
	{
		return m_hwndLV;
	}

private:
	void _UpdateColumns();
	void _UpdateColumnSizes();
	void _UpdateHeaderCheckState(_In_ bool check);

	HWND m_hwndLV = nullptr;
};

class CSmartRenameProgressUI 
	: public IUnknown
{
public:
	CSmartRenameProgressUI(void) noexcept = default;
	virtual ~CSmartRenameProgressUI(void) noexcept = default;

	CSmartRenameProgressUI(CSmartRenameProgressUI const&) = delete;
	CSmartRenameProgressUI& operator=(CSmartRenameProgressUI const&) = delete;

	// IUnknown
	IFACEMETHODIMP QueryInterface(_In_ REFIID riid,_COM_Outptr_ void** ppv)
	{
		static const QITAB qit[] = {
			QITABENT(CSmartRenameProgressUI, IUnknown),
			{0},
		};
		return QISearch(this, qit, riid, ppv);
	}
	IFACEMETHODIMP_(ULONG) AddRef(void)
	{
		return ++m_refCount;
	}
	IFACEMETHODIMP_(ULONG) Release(void)
	{
		auto refCount = --m_refCount;
		if (refCount == 0)
		{
			delete this;
		}
		return refCount;
	}

	HRESULT Start();
	HRESULT Stop();
	bool IsCanceled()
	{
		return m_canceled;
	}

private:
	static LRESULT CALLBACK s_msgWndProc(_In_ HWND hwnd, _In_ uint32_t uMsg, _In_ WPARAM wparam, _In_ LPARAM lparam);
	LRESULT _WndProc(_In_ HWND hwnd, _In_ uint32_t msg, _In_ WPARAM wparam, _In_ LPARAM lparam);

	static uint32_t __stdcall s_workerThread(_In_ void* pv);

	void _UpdateCancelState();
	void _Cleanup();

	wil::unique_process_handle m_workerThreadHandle;
	ATL::CComPtr<IProgressDialog> m_sppd;
	std::atomic<ULONG> m_refCount = 1;
	std::atomic<bool> m_loadingThread {false};
	bool m_canceled = false;
	uint16_t _padding;
};

class CSmartRenameUI 
	: public IDropTarget
	, public ISmartRenameUI
	, public ISmartRenameManagerEvents
{
public:
	CSmartRenameUI(void) noexcept
	{
		m_olecall = wil::OleInitialize();
	}

	// IUnknown
	IFACEMETHODIMP QueryInterface(_In_ REFIID riid,_COM_Outptr_ void** ppv)
	{
		static const QITAB qit[] = {
			QITABENT(CSmartRenameUI, ISmartRenameUI),
			QITABENT(CSmartRenameUI, ISmartRenameManagerEvents),
			QITABENT(CSmartRenameUI, IDropTarget),
			{0},
		};
		return QISearch(this, qit, riid, ppv);
	}
	IFACEMETHODIMP_(ULONG) AddRef(void)
	{
		return ++m_refCount;
	}
	IFACEMETHODIMP_(ULONG) Release(void)
	{
		auto refCount = --m_refCount;
		if (refCount == 0)
		{
			delete this;
		}
		return refCount;
	}

	// ISmartRenameUI
	IFACEMETHODIMP Show(_In_opt_ HWND hwndParent);
	IFACEMETHODIMP Close();
	IFACEMETHODIMP Update();
	IFACEMETHODIMP get_hwnd(_Out_ HWND* hwnd);
	IFACEMETHODIMP get_showUI(_Out_ bool* showUI);

	// ISmartRenameManagerEvents
	IFACEMETHODIMP OnItemAdded(_In_ ISmartRenameItem* renameItem);
	IFACEMETHODIMP OnUpdate(_In_ ISmartRenameItem* renameItem);
	IFACEMETHODIMP OnError(_In_ ISmartRenameItem* renameItem);
	IFACEMETHODIMP OnRegExStarted(_In_ uint32_t threadId);
	IFACEMETHODIMP OnRegExCanceled(_In_ uint32_t threadId);
	IFACEMETHODIMP OnRegExCompleted(_In_ uint32_t threadId);
	IFACEMETHODIMP OnRenameStarted();
	IFACEMETHODIMP OnRenameCompleted();

	// IDropTarget
	IFACEMETHODIMP DragEnter(_In_ IDataObject* pdtobj, ULONG grfKeyState, POINTL pt, _Inout_ ULONG* pdwEffect);
	IFACEMETHODIMP DragOver(ULONG grfKeyState, POINTL pt, _Inout_ ULONG* pdwEffect);
	IFACEMETHODIMP DragLeave();
	IFACEMETHODIMP Drop(_In_ IDataObject* pdtobj, ULONG grfKeyState, POINTL pt, _Inout_ ULONG* pdwEffect);

	static HRESULT s_CreateInstance(_In_ ISmartRenameManager* psrm, _In_opt_ IUnknown* punk, _In_ bool enableDragDrop, _Outptr_ ISmartRenameUI** ppsrui);

private:
	~CSmartRenameUI(void) noexcept
	{
		//DeleteObject(m_iconMain);
		//OleUninitialize();
	}

	CSmartRenameUI(CSmartRenameUI const&) = delete;
	CSmartRenameUI& operator=(CSmartRenameUI const&) = delete;

	HRESULT _DoModal(__in_opt HWND hwnd);
	HRESULT _DoModeless(__in_opt HWND hwnd);

	static intptr_t CALLBACK s_DlgProc(HWND hdlg, uint32_t uMsg, WPARAM wparam, LPARAM lparam) noexcept
	{
		CSmartRenameUI* pDlg = std::bit_cast<CSmartRenameUI*>(::GetWindowLongPtrW(hdlg, DWLP_USER));
		if (uMsg == WM_INITDIALOG)
		{
			pDlg = std::bit_cast<CSmartRenameUI*>(lparam);
			pDlg->m_hwnd = hdlg;
			::SetWindowLongPtrW(hdlg, DWLP_USER, std::bit_cast<intptr_t>(pDlg));
		}
		return pDlg ? pDlg->_DlgProc(uMsg, wparam, lparam) : FALSE;
	}

	HRESULT _Initialize(_In_ ISmartRenameManager* psrm, _In_opt_ IUnknown* punk, _In_ bool enableDragDrop);
	HRESULT _InitAutoComplete();
	void _Cleanup();

	intptr_t _DlgProc(uint32_t uMsg, WPARAM wparam, LPARAM lparam);
	void _OnCommand(_In_ WPARAM wparam, _In_ LPARAM lparam);
	BOOL _OnNotify(_In_ WPARAM wparam, _In_ LPARAM lparam);
	void _OnSize(_In_ WPARAM wparam);
	void _OnGetMinMaxInfo(_In_ LPARAM lparam);
	void _OnInitDlg();
	void _OnRename();
	void _OnAbout();
	void _OnCloseDlg();
	void _OnDestroyDlg();
	void _OnSearchReplaceChanged();
	void _MoveControl(_In_ uint32_t id, _In_ uint32_t repositionFlags, _In_ int xDelta, _In_ int yDelta);

	HRESULT _ReadSettings();
	HRESULT _WriteSettings();

	uint32_t _GetFlagsFromCheckboxes();
	void _SetCheckboxesFromFlags(_In_ uint32_t flags);
	void _ValidateFlagCheckbox(_In_ uint32_t checkBoxId);

	HRESULT _EnumerateItems(_In_ IUnknown* punk);
	void _UpdateCounts();

	CSmartRenameProgressUI m_srpui;
	ATL::CComPtr<ISmartRenameManager> m_spsrm;
	ATL::CComPtr<ISmartRenameEnum> m_spsre;
	ATL::CComPtr<IUnknown> m_spunk;
	ATL::CComPtr<IDropTargetHelper> m_spdth;
	ATL::CComPtr<IAutoComplete2> m_spSearchAC;
	ATL::CComPtr<IUnknown> m_spSearchACL;
	ATL::CComPtr<IAutoComplete2> m_spReplaceAC;
	ATL::CComPtr<IUnknown> m_spReplaceACL;
	CSmartRenameListView m_listview;

	HWND m_hwnd = nullptr;
	HWND m_hwndLV = nullptr;
	wil::unique_hicon m_iconMain;
	uint32_t m_cookie = 0;
	uint32_t m_currentRegExId = 0;
	uint32_t m_selectedCount = 0;
	uint32_t m_renamingCount = 0;
	int32_t m_initialWidth = 0;
	int32_t m_initialHeight = 0;
	int32_t m_lastWidth = 0;
	int32_t m_lastHeight = 0;
	std::atomic<ULONG> m_refCount = 1;
	bool m_initialized = false;
	bool m_enableDragDrop = false;
	bool m_disableCountUpdate = false;
	bool m_modeless = true;
	wil::unique_oleuninitialize_call m_olecall;
	bool _padding[3];
};