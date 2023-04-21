#include "smartrename_pch.h"
#include "common.h"

#include "smartrenameinterfaces.h"
#include "srwlock.h"
#include "smartrenameext.h"
#include "smartrenameui.h"
#include "smartrenameitem.h"
#include "smartrenamemanager.h"
#include "helpers.h"
#include "settings.h"
#include "resource.h"

extern HINSTANCE g_hInst;

struct InvokeStruct
{
	HWND hwndParent;
	IStream* pstrm;
};

HRESULT CSmartRenameMenu::s_CreateInstance(_In_opt_ IUnknown*, _In_ REFIID riid, _Outptr_ void** ppv)
{
	*ppv = nullptr;
	HRESULT hr = E_OUTOFMEMORY;
	CSmartRenameMenu* pprm = new CSmartRenameMenu();
	if (pprm)
	{
		hr = pprm->QueryInterface(riid, ppv);
		pprm->Release();
	}
	return hr;
}

// IShellExtInit
HRESULT CSmartRenameMenu::Initialize(_In_opt_ PCIDLIST_ABSOLUTE, _In_ IDataObject* pdtobj, HKEY)
{
	// Check if we have disabled ourselves
	if (!CSettings::GetEnabled())
		return E_FAIL;

	// Cache the data object to be used later
	helpers::GetShellItemArrayFromUnknown(pdtobj, &m_spia);
	return S_OK;
}

// IContextMenu
HRESULT CSmartRenameMenu::QueryContextMenu(HMENU hMenu, uint32_t index, uint32_t uIDFirst, uint32_t, uint32_t uFlags)
{
	// Check if we have disabled ourselves
	if (!CSettings::GetEnabled())
		return E_FAIL;

	// Check if we should only be on the extended context menu
	if (CSettings::GetExtendedContextMenuOnly() && (!(uFlags & CMF_EXTENDEDVERBS)))
		return E_FAIL;

	HRESULT hr = E_UNEXPECTED;
	if (m_spia && !(uFlags & (CMF_DEFAULTONLY | CMF_VERBSONLY | CMF_OPTIMIZEFORINVOKE)))
	{
		std::wstring_view menuname { helpers::LoadResourceString(IDS_SMARTRENAME) };
		MENUITEMINFOW mii {sizeof(mii)};
		mii.fMask = MIIM_STRING | MIIM_FTYPE | MIIM_ID | MIIM_STATE;
		mii.wID = uIDFirst++;
		mii.fType = MFT_STRING;
		mii.dwTypeData = const_cast<PWSTR>(menuname.data());
		mii.fState = MFS_ENABLED;

		if (CSettings::GetShowIconOnMenu())
		{
			wil::unique_hicon hIcon { static_cast<HICON>(::LoadImageW(g_hInst, util::MakeIntResourceW(IDI_RENAME), IMAGE_ICON, 16, 16, 0)) };
			if (hIcon)
			{
				mii.fMask |= MIIM_BITMAP;
				if (m_hbmpIcon == nullptr)
				{
					m_hbmpIcon.reset(helpers::CreateBitmapFromIcon(hIcon.get()));
				}
				mii.hbmpItem = m_hbmpIcon.get();
			}
		}

		if (!::InsertMenuItemW(hMenu, index, TRUE, &mii))
		{
			hr = HRESULT_FROM_WIN32(GetLastError());
		}
		else
		{
			hr = MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 1);
		}
	}

	return hr;
}

HRESULT CSmartRenameMenu::InvokeCommand(_In_ LPCMINVOKECOMMANDINFO pici)
{
	HRESULT hr = E_FAIL;

	if (CSettings::GetEnabled() && (util::IsIntResource(pici->lpVerb)) /*&& (util::LoWord(pici->lpVerb) == 0)*/)
	{
		hr = _InvokeInternal(pici->hwnd);
	}

	return hr;
}

// IExplorerCommand
IFACEMETHODIMP CSmartRenameMenu::GetTitle([[maybe_unused]] _In_opt_ IShellItemArray* psia, _Outptr_result_nullonfailure_ PWSTR* name)
{
	*name = nullptr;
	std::wstring_view menuname { helpers::LoadResourceString(IDS_SMARTRENAME) };

	return ::SHStrDupW(menuname.data(), name);
}

IFACEMETHODIMP CSmartRenameMenu::GetIcon(_In_opt_ IShellItemArray*, _Outptr_result_nullonfailure_ PWSTR* icon)
{
	*icon = nullptr;
#ifdef _WIN64
	return ::SHStrDupW(L"SmartRenameExt64.dll,-132", icon);
#else
	return ::SHStrDupW(L"SmartRenameExt32.dll,-132", icon);
#endif
}

IFACEMETHODIMP CSmartRenameMenu::GetState([[maybe_unused]] _In_opt_ IShellItemArray* psia,
	[[maybe_unused]] _In_ BOOL okToBeSlow, _Out_ EXPCMDSTATE* cmdState)
{
	// Check if we have disabled ourselves
	*cmdState = (CSettings::GetEnabled()) ? ECS_ENABLED : ECS_HIDDEN;
	return S_OK;
}

IFACEMETHODIMP CSmartRenameMenu::Invoke(_In_opt_ IShellItemArray* psia, _In_opt_ IBindCtx*)
{
	HRESULT hr = E_FAIL;
	if (CSettings::GetEnabled() && psia)
	{
		m_spia = psia;
		hr = _InvokeInternal(nullptr);
	}

	return hr;
}

// IObjectWithSites
IFACEMETHODIMP CSmartRenameMenu::SetSite(_In_ IUnknown* punk)
{
	m_spSite = punk;
	return S_OK;
}

IFACEMETHODIMP CSmartRenameMenu::GetSite(_In_ REFIID riid, _COM_Outptr_ void** ppunk)
{
	*ppunk = nullptr;
	HRESULT hr = E_FAIL;
	if (m_spSite)
	{
		hr = m_spSite->QueryInterface(riid, ppunk);
	}
	return hr;
}

HRESULT CSmartRenameMenu::_InvokeInternal(_In_opt_ HWND hwndParent) noexcept
{
	InvokeStruct* pInvokeData = new InvokeStruct;
	HRESULT hr = pInvokeData ? S_OK : E_OUTOFMEMORY;
	if (SUCCEEDED(hr))
	{
		pInvokeData->hwndParent = hwndParent;
		hr = CoMarshalInterThreadInterfaceInStream(__uuidof(m_spia), m_spia, &(pInvokeData->pstrm));
		if (SUCCEEDED(hr))
		{
			hr = SHCreateThread(s_SmartRenameUIThreadProc, pInvokeData, CTF_COINIT | CTF_PROCESS_REF, nullptr) ? S_OK : E_FAIL;
			if (FAILED(hr))
			{
				pInvokeData->pstrm->Release(); // if we failed to create the thread, then we must release the stream
			}
		}

		if (FAILED(hr))
		{
			delete pInvokeData;
		}
	}

	return hr;
}

ULONG __stdcall CSmartRenameMenu::s_SmartRenameUIThreadProc(_In_ void* pData) noexcept
{
	InvokeStruct* pInvokeData = static_cast<InvokeStruct*>(pData);
	ATL::CComPtr<IUnknown> spunk;
	if (SUCCEEDED(CoGetInterfaceAndReleaseStream(pInvokeData->pstrm, IID_PPV_ARGS(&spunk))))
	{
		// Create the smart rename manager
		ATL::CComPtr<ISmartRenameManager> spsrm;
		if (SUCCEEDED(CSmartRenameManager::s_CreateInstance(&spsrm)))
		{
			// Create the factory for our items
			ATL::CComPtr<ISmartRenameItemFactory> spsrif;
			if (SUCCEEDED(CSmartRenameItem::s_CreateInstance(nullptr, IID_PPV_ARGS(&spsrif))))
			{
				// Pass the factory to the manager
				if (SUCCEEDED(spsrm->put_renameItemFactory(spsrif)))
				{
					// Create the smart rename UI instance and pass the smart rename manager
					ATL::CComPtr<ISmartRenameUI> spsrui;
					if (SUCCEEDED(CSmartRenameUI::s_CreateInstance(spsrm, spunk, false, &spsrui)))
					{
						// Call blocks until we are done
						spsrui->Show(pInvokeData->hwndParent);
						spsrui->Close();
					}
				}
			}

			// Need to call shutdown to break circular dependencies
			spsrm->Shutdown();
		}
	}

	delete pInvokeData;

	return 0;
}
