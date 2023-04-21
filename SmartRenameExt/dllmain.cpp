#include "smartrename_pch.h"
#include "common.h"

#include "SmartRenameExt.h"

std::atomic<uint32_t> g_dwModuleRefCount = 0;
HINSTANCE g_hInst = nullptr;

class CSmartRenameClassFactory : public IClassFactory
{
public:
	CSmartRenameClassFactory(_In_ REFCLSID clsid) noexcept
		: m_clsid(clsid)
	{
		DllAddRef();
	}

	CSmartRenameClassFactory(CSmartRenameClassFactory const&) = delete;
	CSmartRenameClassFactory& operator=(CSmartRenameClassFactory const&) = delete;

	// IUnknown methods
	IFACEMETHODIMP QueryInterface(_In_ REFIID riid, _COM_Outptr_ void** ppv)
	{
		static const QITAB qit[] = {
			QITABENT(CSmartRenameClassFactory, IClassFactory),
			{0}};
		return QISearch(this, qit, riid, ppv);
	}

	IFACEMETHODIMP_(ULONG)
	AddRef(void)
	{
		return static_cast<ULONG>(++m_refCount);
	}

	IFACEMETHODIMP_(ULONG)
	Release(void)
	{
		auto refCount = --m_refCount;
		if (refCount == 0)
		{
			delete this;
		}
		return static_cast<ULONG>(refCount);
	}

	// IClassFactory methods
	IFACEMETHODIMP CreateInstance(_In_opt_ IUnknown* punkOuter, _In_ REFIID riid, _Outptr_ void** ppv)
	{
		*ppv = nullptr;
		HRESULT hr;
		if (punkOuter)
		{
			hr = CLASS_E_NOAGGREGATION;
		}
		else
		{
			if (m_clsid == CLSID_SmartRenameMenu)
			{
				hr = CSmartRenameMenu::s_CreateInstance(punkOuter, riid, ppv);
			}
			else
			{
				hr = CLASS_E_CLASSNOTAVAILABLE;
			}
		}
		return hr;
	}

	IFACEMETHODIMP LockServer(BOOL bLock)
	{
		if (bLock)
		{
			DllAddRef();
		}
		else
		{
			DllRelease();
		}
		return S_OK;
	}

private:
	~CSmartRenameClassFactory()
	{
		DllRelease();
	}

	std::atomic<long> m_refCount = 1;
	CLSID m_clsid;
};

/// <summary>
/// <c>DllMain</c>
/// </summary>
/// <remarks>
/// </remarks>
/// <param name="dllhandle"></param>
/// <param name="reason"></param>
/// <param name="reserved"></param>
/// <returns>BOOL</returns>
extern "C" BOOL __stdcall DllMain([[maybe_unused]] _In_ HINSTANCE dllhandle, [[maybe_unused]] _In_ uint32_t reason, [[maybe_unused]] _In_opt_ void* reserved)
{
	switch (reason)
	{
	case DLL_PROCESS_ATTACH:
		THROW_IF_WIN32_BOOL_FALSE(::DisableThreadLibraryCalls(dllhandle));
		g_hInst = dllhandle;
		break;
	case DLL_THREAD_ATTACH:
		break;
	case DLL_THREAD_DETACH:
		break;
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}

//
// Checks if there are any external references to this module
//
STDAPI DllCanUnloadNow(void)
{
	return (g_dwModuleRefCount == 0) ? S_OK : S_FALSE;
}

//
// DLL export for creating COM objects
//
STDAPI DllGetClassObject(_In_ REFCLSID clsid, _In_ REFIID riid, _Outptr_ void** ppv)
{
	*ppv = nullptr;
	HRESULT hr = E_OUTOFMEMORY;
	CSmartRenameClassFactory* pClassFactory = new CSmartRenameClassFactory(clsid);
	if (pClassFactory)
	{
		hr = pClassFactory->QueryInterface(riid, ppv);
		pClassFactory->Release();
	}
	return hr;
}

STDAPI DllRegisterServer(void)
{
	return S_OK;
}

STDAPI DllUnregisterServer(void)
{
	return S_OK;
}

void DllAddRef()
{
	g_dwModuleRefCount++;
}

void DllRelease()
{
	g_dwModuleRefCount--;
}
