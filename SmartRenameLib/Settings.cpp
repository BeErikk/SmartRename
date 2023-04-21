#include "smartrename_pch.h"
#include "common.h"

#include "settings.h"
#include "smartrenameinterfaces.h"

//#include <commctrl.h>
namespace literals
{
constexpr std::wstring_view c_rootRegPath {L"Software\\SmartRename"};
constexpr std::wstring_view c_mruSearchRegPath {L"SearchMRU"};
constexpr std::wstring_view c_mruReplaceRegPath {L"ReplaceMRU"};
constexpr std::wstring_view c_enabled {L"Enabled"};
constexpr std::wstring_view c_showIconOnMenu {L"ShowIcon"};
constexpr std::wstring_view c_extendedContextMenuOnly {L"ExtendedContextMenuOnly"};
constexpr std::wstring_view c_persistState {L"PersistState"};
constexpr std::wstring_view c_maxMRUSize {L"MaxMRUSize"};
constexpr std::wstring_view c_flags {L"Flags"};
constexpr std::wstring_view c_searchText {L"SearchText"};
constexpr std::wstring_view c_replaceText {L"::ReplaceTextW"};
constexpr std::wstring_view c_mruEnabled {L"MRUEnabled"};
} // namespace literals

const bool c_enabledDefault = true;
const bool c_showIconOnMenuDefault = true;
const bool c_extendedContextMenuOnlyDefaut = false;
const bool c_persistStateDefault = true;
const bool c_mruEnabledDefault = true;

const uint32_t c_maxMRUSizeDefault = 10;
const uint32_t c_flagsDefault = 0;

bool CSettings::GetEnabled()
{
	return GetRegBoolValue(literals::c_enabled.data(), c_enabledDefault);
}

bool CSettings::SetEnabled(_In_ bool enabled)
{
	return SetRegBoolValue(literals::c_enabled.data(), enabled);
}

bool CSettings::GetShowIconOnMenu()
{
	return GetRegBoolValue(literals::c_showIconOnMenu.data(), c_showIconOnMenuDefault);
}

bool CSettings::SetShowIconOnMenu(_In_ bool show)
{
	return SetRegBoolValue(literals::c_showIconOnMenu.data(), show);
}

bool CSettings::GetExtendedContextMenuOnly()
{
	return GetRegBoolValue(literals::c_extendedContextMenuOnly.data(), c_extendedContextMenuOnlyDefaut);
}

bool CSettings::SetExtendedContextMenuOnly(_In_ bool extendedOnly)
{
	return SetRegBoolValue(literals::c_extendedContextMenuOnly.data(), extendedOnly);
}

bool CSettings::GetPersistState()
{
	return GetRegBoolValue(literals::c_persistState.data(), c_persistStateDefault);
}

bool CSettings::SetPersistState(_In_ bool persistState)
{
	return SetRegBoolValue(literals::c_persistState.data(), persistState);
}

bool CSettings::GetMRUEnabled()
{
	return GetRegBoolValue(literals::c_mruEnabled.data(), c_mruEnabledDefault);
}

bool CSettings::SetMRUEnabled(_In_ bool enabled)
{
	return SetRegBoolValue(literals::c_mruEnabled.data(), enabled);
}

uint32_t CSettings::GetMaxMRUSize()
{
	return GetRegDWORDValue(literals::c_maxMRUSize.data(), c_maxMRUSizeDefault);
}

bool CSettings::SetMaxMRUSize(_In_ uint32_t maxMRUSize)
{
	return SetRegDWORDValue(literals::c_maxMRUSize.data(), maxMRUSize);
}

uint32_t CSettings::GetFlags()
{
	return GetRegDWORDValue(literals::c_flags.data(), c_flagsDefault);
}

bool CSettings::SetFlags(_In_ uint32_t flags)
{
	return SetRegDWORDValue(literals::c_flags.data(), flags);
}

bool CSettings::GetSearchText(__out_ecount(cchBuf) PWSTR text, uint32_t cchBuf)
{
	return GetRegStringValue(literals::c_searchText.data(), text, cchBuf);
}

bool CSettings::SetSearchText(_In_ PCWSTR text)
{
	return SetRegStringValue(literals::c_searchText.data(), text);
}

bool CSettings::GetReplaceText(__out_ecount(cchBuf) PWSTR text, uint32_t cchBuf)
{
	return GetRegStringValue(literals::c_replaceText.data(), text, cchBuf);
}

bool CSettings::SetReplaceText(_In_ PCWSTR text)
{
	return SetRegStringValue(literals::c_replaceText.data(), text);
}

bool CSettings::SetRegBoolValue(_In_ PCWSTR valueName, _In_ bool value)
{
	uint32_t dwValue = value ? 1 : 0;
	return SetRegDWORDValue(valueName, dwValue);
}

bool CSettings::GetRegBoolValue(_In_ PCWSTR valueName, _In_ bool defaultValue)
{
	uint32_t value = GetRegDWORDValue(valueName, (defaultValue == 0) ? false : true);
	return (value == 0) ? false : true;
}

bool CSettings::SetRegDWORDValue(_In_ PCWSTR valueName, _In_ uint32_t value)
{
	return (SUCCEEDED(HRESULT_FROM_WIN32(SHSetValue(HKEY_CURRENT_USER, literals::c_rootRegPath.data(), valueName, REG_DWORD, &value, sizeof(value)))));
}

uint32_t CSettings::GetRegDWORDValue(_In_ PCWSTR valueName, _In_ uint32_t defaultValue)
{
	uint32_t retVal = defaultValue;
	uint32_t type = REG_DWORD;
	uint32_t dwEnabled = 0;
	uint32_t cb = sizeof(dwEnabled);
	if (::SHGetValueW(HKEY_CURRENT_USER, literals::c_rootRegPath.data(), valueName, &type, &dwEnabled, &cb) == ERROR_SUCCESS)
	{
		retVal = dwEnabled;
	}

	return retVal;
}

bool CSettings::SetRegStringValue(_In_ PCWSTR valueName, _In_ PCWSTR value)
{
	ULONG cb = (uint32_t)((wcslen(value) + 1) * sizeof(*value));
	return (SUCCEEDED(HRESULT_FROM_WIN32(SHSetValue(HKEY_CURRENT_USER, literals::c_rootRegPath.data(), valueName, REG_SZ, (const BYTE*)value, cb))));
}

bool CSettings::GetRegStringValue(_In_ PCWSTR valueName, __out_ecount(cchBuf) PWSTR value, uint32_t cchBuf)
{
	if (cchBuf > 0)
	{
		value[0] = L'\0';
	}

	uint32_t type = REG_SZ;
	ULONG cb = cchBuf * sizeof(*value);
	return (SUCCEEDED(HRESULT_FROM_WIN32(::SHGetValueW(HKEY_CURRENT_USER, literals::c_rootRegPath.data(), valueName, &type, value, &cb) == ERROR_SUCCESS)));
}

typedef int(CALLBACK* MRUCMPPROC)(LPCWSTR, LPCWSTR);

typedef struct
{
	uint32_t cbSize;
	uint32_t uMax;
	uint32_t fFlags;
	HKEY hKey;
	PCWSTR lpszSubKey;
	MRUCMPPROC lpfnCompare;
} MRUINFO;

typedef HANDLE(__stdcall* CreateMRUListFn)(MRUINFO* pmi);
typedef int(__stdcall* AddMRUStringFn)(HANDLE hMRU, LPCWSTR data);
typedef int(__stdcall* EnumMRUListFn)(HANDLE hMRU, int nItem, void* lpData, uint32_t uLen);
typedef int(__stdcall* FreeMRUListFn)(HANDLE hMRU);

class CRenameMRU : public IEnumString,
				   public ISmartRenameMRU
{
public:
	// IUnknown
	IFACEMETHODIMP_(ULONG)
	AddRef();
	IFACEMETHODIMP_(ULONG)
	Release();
	IFACEMETHODIMP QueryInterface(_In_ REFIID riid, _Outptr_ void** ppv);

	// IEnumString
	IFACEMETHODIMP Next(__in ULONG celt, __out_ecount_part(celt, *pceltFetched) LPOLESTR* rgelt, __out_opt ULONG* pceltFetched);
	IFACEMETHODIMP Skip(__in ULONG /*celt*/)
	{
		return E_NOTIMPL;
	}
	IFACEMETHODIMP Reset();
	IFACEMETHODIMP Clone(__deref_out IEnumString** ppenum)
	{
		*ppenum = nullptr;
		return E_NOTIMPL;
	}

	// ISmartRenameMRU
	IFACEMETHODIMP AddMRUString(_In_ PCWSTR entry);

	static HRESULT CreateInstance(_In_ PCWSTR regPathMRU, _In_ ULONG maxMRUSize, _Outptr_ IUnknown** ppUnk);

private:
	CRenameMRU();
	~CRenameMRU();

	HRESULT _Initialize(_In_ PCWSTR regPathMRU, __in ULONG maxMRUSize);
	HRESULT _CreateMRUList(_In_ MRUINFO* pmi);
	int _AddMRUString(_In_ PCWSTR data);
	int _EnumMRUList(_In_ int nItem, _Out_ void* lpData, _In_ uint32_t uLen);
	void _FreeMRUList();

	std::atomic<ULONG> m_refCount = 1;
	HKEY m_hKey = nullptr;
	ULONG m_maxMRUSize = 0;
	ULONG m_mruIndex = 0;
	ULONG m_mruSize = 0;
	HANDLE m_mruHandle = nullptr;
	HMODULE m_hComctl32Dll = nullptr;
	wil::unique_cotaskmem_string m_regPath;
};

CRenameMRU::CRenameMRU()
	: m_refCount(1)
{
}

CRenameMRU::~CRenameMRU()
{
	if (m_hKey)
	{
		RegCloseKey(m_hKey);
	}

	_FreeMRUList();

	if (m_hComctl32Dll)
	{
		FreeLibrary(m_hComctl32Dll);
	}

	//CoTaskMemFree(m_regPath);
}

HRESULT CRenameMRU::CreateInstance(_In_ PCWSTR regPathMRU, _In_ ULONG maxMRUSize, _Outptr_ IUnknown** ppUnk)
{
	*ppUnk = nullptr;
	HRESULT hr = (regPathMRU && maxMRUSize > 0) ? S_OK : E_FAIL;
	if (SUCCEEDED(hr))
	{
		CRenameMRU* renameMRU = new CRenameMRU();
		hr = renameMRU ? S_OK : E_OUTOFMEMORY;
		if (SUCCEEDED(hr))
		{
			hr = renameMRU->_Initialize(regPathMRU, maxMRUSize);
			if (SUCCEEDED(hr))
			{
				hr = renameMRU->QueryInterface(IID_PPV_ARGS(ppUnk));
			}

			renameMRU->Release();
		}
	}

	return hr;
}

// IUnknown
IFACEMETHODIMP_(ULONG)
CRenameMRU::AddRef()
{
	return InterlockedIncrement(&m_refCount);
}

IFACEMETHODIMP_(ULONG)
CRenameMRU::Release()
{
	long refCount = InterlockedDecrement(&m_refCount);

	if (refCount == 0)
	{
		delete this;
	}
	return refCount;
}

IFACEMETHODIMP CRenameMRU::QueryInterface(_In_ REFIID riid, _Outptr_ void** ppv)
{
	static const QITAB qit[] = {
		QITABENT(CRenameMRU, IEnumString),
		QITABENT(CRenameMRU, ISmartRenameMRU),
		{0}};
	return QISearch(this, qit, riid, ppv);
}

HRESULT CRenameMRU::_Initialize(_In_ PCWSTR regPathMRU, __in ULONG maxMRUSize)
{
	m_maxMRUSize = maxMRUSize;

	wchar_t regPath[MAX_PATH] = {0};
	HRESULT hr = StringCchPrintf(regPath, std::size(regPath), L"%s\\%s", literals::c_rootRegPath.data(), regPathMRU);
	if (SUCCEEDED(hr))
	{
		hr = ::SHStrDupW(regPathMRU, &m_regPath);

		if (SUCCEEDED(hr))
		{
			MRUINFO mi = {
				sizeof(MRUINFO),
				maxMRUSize,
				0,
				HKEY_CURRENT_USER,
				regPath,
				nullptr};

			hr = _CreateMRUList(&mi);
			if (SUCCEEDED(hr))
			{
				m_mruSize = _EnumMRUList(-1, nullptr, 0);
			}
			else
			{
				hr = E_FAIL;
			}
		}
	}

	return hr;
}

// IEnumString
IFACEMETHODIMP CRenameMRU::Reset()
{
	m_mruIndex = 0;
	return S_OK;
}

#define MAX_ENTRY_STRING 1024

IFACEMETHODIMP CRenameMRU::Next(__in ULONG celt, __out_ecount_part(celt, *pceltFetched) LPOLESTR* rgelt, __out_opt ULONG* pceltFetched)
{
	HRESULT hr = S_OK;
	WCHAR mruEntry[MAX_ENTRY_STRING];
	mruEntry[0] = L'\0';

	if (pceltFetched)
	{
		*pceltFetched = 0;
	}

	if (!celt)
	{
		return S_OK;
	}

	if (!rgelt)
	{
		return S_FALSE;
	}

	hr = S_FALSE;
	if (m_mruIndex <= m_mruSize && _EnumMRUList(m_mruIndex++, (void*)mruEntry, std::size(mruEntry)) > 0)
	{
		hr = ::SHStrDupW(mruEntry, rgelt);
		if (SUCCEEDED(hr) && pceltFetched != nullptr)
		{
			*pceltFetched = 1;
		}
	}

	return hr;
}

IFACEMETHODIMP CRenameMRU::AddMRUString(_In_ PCWSTR entry)
{
	return (_AddMRUString(entry) < 0) ? E_FAIL : S_OK;
}

HRESULT CRenameMRU::_CreateMRUList(_In_ MRUINFO* pmi)
{
	if (m_mruHandle != nullptr)
	{
		_FreeMRUList();
	}

	if (m_hComctl32Dll == nullptr)
	{
		m_hComctl32Dll = LoadLibraryEx(L"comctl32.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
	}

	if (m_hComctl32Dll != nullptr)
	{
		CreateMRUListFn pfnCreateMRUList = std::bit_cast<CreateMRUListFn>(GetProcAddress(m_hComctl32Dll, util::MakeIntResource<char>(400)));
		if (pfnCreateMRUList != nullptr)
		{
			m_mruHandle = pfnCreateMRUList(pmi);
		}
	}

	return (m_mruHandle != nullptr) ? S_OK : E_FAIL;
}

int CRenameMRU::_AddMRUString(_In_ PCWSTR data)
{
	int retVal = -1;
	if (m_mruHandle != nullptr)
	{
		if (m_hComctl32Dll == nullptr)
		{
			m_hComctl32Dll = LoadLibraryEx(L"comctl32.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
		}

		if (m_hComctl32Dll != nullptr)
		{
			AddMRUStringFn pfnAddMRUString = std::bit_cast<AddMRUStringFn>(GetProcAddress(m_hComctl32Dll, util::MakeIntResource<char>(401)));
			if (pfnAddMRUString != nullptr)
			{
				retVal = pfnAddMRUString(m_mruHandle, data);
			}
		}
	}

	return retVal;
}

int CRenameMRU::_EnumMRUList(_In_ int nItem, _Out_ void* lpData, _In_ uint32_t uLen)
{
	lpData = nullptr;
	int retVal = -1;
	if (m_mruHandle != nullptr)
	{
		if (m_hComctl32Dll == nullptr)
		{
			m_hComctl32Dll = LoadLibraryEx(L"comctl32.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
		}

		if (m_hComctl32Dll != nullptr)
		{
			EnumMRUListFn pfnEnumMRUList = std::bit_cast<EnumMRUListFn>(GetProcAddress(m_hComctl32Dll, util::MakeIntResource<char>(403)));
			if (pfnEnumMRUList != nullptr)
			{
				retVal = pfnEnumMRUList(m_mruHandle, nItem, lpData, uLen);
			}
		}
	}

	return retVal;
}

void CRenameMRU::_FreeMRUList()
{
	if (m_mruHandle != nullptr)
	{
		if (m_hComctl32Dll == nullptr)
		{
			m_hComctl32Dll = LoadLibraryEx(L"comctl32.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
		}

		if (m_hComctl32Dll != nullptr)
		{
			FreeMRUListFn pfnFreeMRUList = std::bit_cast<FreeMRUListFn>(GetProcAddress(m_hComctl32Dll, util::MakeIntResource<char>(152)));
			if (pfnFreeMRUList != nullptr)
			{
				pfnFreeMRUList(m_mruHandle);
			}
		}
		m_mruHandle = nullptr;
	}
}

HRESULT CRenameMRUSearch_CreateInstance(_Outptr_ IUnknown** ppUnk)
{
	return CRenameMRU::CreateInstance(literals::c_mruSearchRegPath.data(), CSettings::GetMaxMRUSize(), ppUnk);
}

HRESULT CRenameMRUReplace_CreateInstance(_Outptr_ IUnknown** ppUnk)
{
	return CRenameMRU::CreateInstance(literals::c_mruReplaceRegPath.data(), CSettings::GetMaxMRUSize(), ppUnk);
}
