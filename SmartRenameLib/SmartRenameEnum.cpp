#include "smartrename_pch.h"
#include "common.h"

#include "smartrenameinterfaces.h"
#include "srwlock.h"
#include "SmartRenameEnum.h"
#include "helpers.h"

//#include <ShlGuid.h>

IFACEMETHODIMP CSmartRenameEnum::Start()
{
	m_canceled = false;
	ATL::CComPtr<IShellItemArray> spsia;
	HRESULT hr = helpers::GetShellItemArrayFromUnknown(m_spunk, &spsia);
	if (SUCCEEDED(hr))
	{
		ATL::CComPtr<IEnumShellItems> spesi;
		hr = spsia->EnumItems(&spesi);
		if (SUCCEEDED(hr))
		{
			hr = _ParseEnumItems(spesi);
		}
	}

	return hr;
}

IFACEMETHODIMP CSmartRenameEnum::Cancel()
{
	m_canceled = true;
	return S_OK;
}

HRESULT CSmartRenameEnum::s_CreateInstance(_In_ IUnknown* punk, _In_ ISmartRenameManager* psrm, _In_ REFIID iid, _Outptr_ void** resultInterface)
{
	*resultInterface = nullptr;

	CSmartRenameEnum* newRenameEnum = new CSmartRenameEnum();
	HRESULT hr = newRenameEnum ? S_OK : E_OUTOFMEMORY;
	if (SUCCEEDED(hr))
	{
		hr = newRenameEnum->_Init(punk, psrm);
		if (SUCCEEDED(hr))
		{
			hr = newRenameEnum->QueryInterface(iid, resultInterface);
		}

		newRenameEnum->Release();
	}
	return hr;
}

HRESULT CSmartRenameEnum::_Init(_In_ IUnknown* punk, _In_ ISmartRenameManager* psrm)
{
	m_spunk = punk;
	m_spsrm = psrm;
	return S_OK;
}

HRESULT CSmartRenameEnum::_ParseEnumItems(_In_ IEnumShellItems* pesi, _In_ int depth)
{
	HRESULT hr = E_INVALIDARG;

	// We shouldn't get this deep since we only enum the contents of
	// regular folders but adding just in case
	if ((pesi) && (depth < (MAX_PATH / 2)))
	{
		hr = S_OK;

		ULONG celtFetched;
		ATL::CComPtr<IShellItem> spsi;
		while ((S_OK == pesi->Next(1, &spsi, &celtFetched)) && (SUCCEEDED(hr)))
		{
			if (m_canceled)
			{
				return E_ABORT;
			}

			ATL::CComPtr<ISmartRenameItemFactory> spsrif;
			hr = m_spsrm->get_renameItemFactory(&spsrif);
			if (SUCCEEDED(hr))
			{
				ATL::CComPtr<ISmartRenameItem> spNewItem;
				// Failure may be valid if we come across a shell item that does
				// not support a file system path.  In that case we simply ignore
				// the item.
				if (SUCCEEDED(spsrif->Create(spsi, &spNewItem)))
				{
					spNewItem->put_depth(depth);
					hr = m_spsrm->AddItem(spNewItem);
					if (SUCCEEDED(hr))
					{
						bool isFolder = false;
						if (SUCCEEDED(spNewItem->get_isFolder(&isFolder)) && isFolder)
						{
							// Bind to the IShellItem for the IEnumShellItems interface
							ATL::CComPtr<IEnumShellItems> spesiNext;
							hr = spsi->BindToHandler(nullptr, BHID_EnumItems, IID_PPV_ARGS(&spesiNext));
							if (SUCCEEDED(hr))
							{
								// Parse the folder contents recursively
								hr = _ParseEnumItems(spesiNext, depth + 1);
							}
						}
					}
				}
			}

			spsi = nullptr;
		}
	}

	return hr;
}
