#pragma once

class CSettings
{
public:
	static bool GetEnabled();
	static bool SetEnabled(_In_ bool enabled);

	static bool GetShowIconOnMenu();
	static bool SetShowIconOnMenu(_In_ bool show);

	static bool GetExtendedContextMenuOnly();
	static bool SetExtendedContextMenuOnly(_In_ bool extendedOnly);

	static bool GetPersistState();
	static bool SetPersistState(_In_ bool extendedOnly);

	static bool GetMRUEnabled();
	static bool SetMRUEnabled(_In_ bool enabled);

	static uint32_t GetMaxMRUSize();
	static bool SetMaxMRUSize(_In_ uint32_t maxMRUSize);

	static uint32_t GetFlags();
	static bool SetFlags(_In_ uint32_t flags);

	static bool GetSearchText(__out_ecount(cchBuf) PWSTR text, uint32_t cchBuf);
	static bool SetSearchText(_In_ PCWSTR text);

	static bool GetReplaceText(__out_ecount(cchBuf) PWSTR text, uint32_t cchBuf);
	static bool SetReplaceText(_In_ PCWSTR text);

private:
	static bool GetRegBoolValue(_In_ PCWSTR valueName, _In_ bool defaultValue);
	static bool SetRegBoolValue(_In_ PCWSTR valueName, _In_ bool value);
	static bool SetRegDWORDValue(_In_ PCWSTR valueName, _In_ uint32_t value);
	static uint32_t GetRegDWORDValue(_In_ PCWSTR valueName, _In_ uint32_t defaultValue);
	static bool SetRegStringValue(_In_ PCWSTR valueName, _In_ PCWSTR value);
	static bool GetRegStringValue(_In_ PCWSTR valueName, __out_ecount(cchBuf) PWSTR value, uint32_t cchBuf);
};

HRESULT CRenameMRUSearch_CreateInstance(_Outptr_ IUnknown** ppUnk);
HRESULT CRenameMRUReplace_CreateInstance(_Outptr_ IUnknown** ppUnk);