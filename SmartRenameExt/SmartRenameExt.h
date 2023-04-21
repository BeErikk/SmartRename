#pragma once
//#include "smartrename_pch.h"

extern void DllAddRef(void);
extern void DllRelease(void);

#define INITGUID
#include <guiddef.h>

// {81ADB5B6-F9A4-4320-87B3-D9360F82EC50}
DEFINE_GUID(CLSID_SmartRenameMenu, 0x81ADB5B6, 0xF9A4, 0x4320, 0x87, 0xB3, 0xD9, 0x36, 0x0F, 0x82, 0xEC, 0x50);

class CSmartRenameMenu : public IShellExtInit,
						 public IContextMenu,
						 public IExplorerCommand,
						 public IObjectWithSite
{
public:
	CSmartRenameMenu(void) noexcept
	{
		DllAddRef();
	}

	CSmartRenameMenu(CSmartRenameMenu const&) = delete;
	CSmartRenameMenu& operator=(CSmartRenameMenu const&) = delete;

	// IUnknown
	IFACEMETHODIMP QueryInterface(_In_ REFIID riid, _COM_Outptr_ void** ppv)
	{
		static const QITAB qit[] = {
			QITABENT(CSmartRenameMenu, IShellExtInit),
			QITABENT(CSmartRenameMenu, IContextMenu),
			QITABENT(CSmartRenameMenu, IExplorerCommand),
			QITABENT(CSmartRenameMenu, IObjectWithSite),
			{0, 0},
		};
		return QISearch(this, qit, riid, ppv);
	}

	IFACEMETHODIMP_(ULONG) AddRef()
	{
		return static_cast<ULONG>(++m_refCount);
	}

	IFACEMETHODIMP_(ULONG) Release()
	{
		auto refCount = --m_refCount;
		if (refCount == 0)
		{
			delete this;
		}
		return static_cast<ULONG>(refCount);
	}

	// IShellExtInit
	STDMETHODIMP Initialize(_In_opt_ PCIDLIST_ABSOLUTE pidlFolder, _In_ IDataObject* pdto, HKEY hkProgID);

	// IContextMenu
	STDMETHODIMP QueryContextMenu(HMENU hMenu, uint32_t index, uint32_t uIDFirst, uint32_t uIDLast, uint32_t uFlags);
	STDMETHODIMP InvokeCommand(_In_ LPCMINVOKECOMMANDINFO pCMI);
	STDMETHODIMP GetCommandString(uintptr_t, uint32_t, _In_opt_ uint32_t*, _In_ LPSTR, uint32_t)
	{
		return E_NOTIMPL;
	}

	// IExplorerCommand
	IFACEMETHODIMP GetTitle(_In_opt_ IShellItemArray* psia, _Outptr_result_nullonfailure_ PWSTR* name);
	IFACEMETHODIMP GetIcon(_In_opt_ IShellItemArray*, _Outptr_result_nullonfailure_ PWSTR* icon);
	IFACEMETHODIMP GetToolTip(_In_opt_ IShellItemArray*, _Outptr_result_nullonfailure_ PWSTR* infoTip)
	{
		*infoTip = nullptr;
		return E_NOTIMPL;
	}
	IFACEMETHODIMP GetCanonicalName(_Out_ GUID* guidCommandName)
	{
		*guidCommandName = GUID_NULL;
		return S_OK;
	}
	IFACEMETHODIMP GetState(_In_opt_ IShellItemArray* psia, _In_ BOOL okToBeSlow, _Out_ EXPCMDSTATE* cmdState);
	IFACEMETHODIMP Invoke(_In_opt_ IShellItemArray* psia, _In_opt_ IBindCtx*);
	IFACEMETHODIMP GetFlags(_Out_ EXPCMDFLAGS* flags)
	{
		*flags = ECF_DEFAULT;
		return S_OK;
	}
	IFACEMETHODIMP EnumSubCommands(_COM_Outptr_ IEnumExplorerCommand** ppeec)
	{
		*ppeec = nullptr;
		return E_NOTIMPL;
	}

	// IObjectWithSite
	IFACEMETHODIMP SetSite(_In_ IUnknown* pink);
	IFACEMETHODIMP GetSite(_In_ REFIID riid, _COM_Outptr_ void** ppvoid);

	static HRESULT s_CreateInstance(_In_opt_ IUnknown* punkOuter, _In_ REFIID riid, _Outptr_ void** ppv);
	static ULONG __stdcall s_SmartRenameUIThreadProc(_In_ void* pData) noexcept;

private:
	~CSmartRenameMenu(void) noexcept
	{
		//m_spia = nullptr;
		//DeleteObject(m_hbmpIcon);
		DllRelease();
	}

	HRESULT _InvokeInternal(_In_opt_ HWND hwndParent) noexcept;

	wil::unique_hbitmap m_hbmpIcon;
	ATL::CComPtr<IShellItemArray> m_spia;
	ATL::CComPtr<IUnknown> m_spSite;
	std::atomic<long> m_refCount = 1;
};
