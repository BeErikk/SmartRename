#pragma once

//#include "smartrename_pch.h"
//#include <vector>
//#include <string>
//#include "srwlock.h"

#define DEFAULT_FLAGS MatchAllOccurrences

class CSmartRenameRegEx : public ISmartRenameRegEx
{
public:
	// IUnknown
	IFACEMETHODIMP QueryInterface(_In_ REFIID riid,_COM_Outptr_ void** ppv)
	{
		static const QITAB qit[] = {
			QITABENT(CSmartRenameRegEx, ISmartRenameRegEx),
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

	// ISmartRenameRegEx
	IFACEMETHODIMP Advise(_In_ ISmartRenameRegExEvents* regExEvents, _Out_ uint32_t* cookie);
	IFACEMETHODIMP UnAdvise(_In_ uint32_t cookie);
	IFACEMETHODIMP get_searchTerm(_Outptr_ PWSTR* searchTerm);
	IFACEMETHODIMP put_searchTerm(_In_ PCWSTR searchTerm);
	IFACEMETHODIMP get_replaceTerm(_Outptr_ PWSTR* replaceTerm);
	IFACEMETHODIMP put_replaceTerm(_In_ PCWSTR replaceTerm);
	IFACEMETHODIMP get_flags(_Out_ uint32_t* flags);
	IFACEMETHODIMP put_flags(_In_ uint32_t flags);
	IFACEMETHODIMP Replace(_In_ PCWSTR source, _Outptr_ PWSTR* result);

	static HRESULT s_CreateInstance(_Outptr_ ISmartRenameRegEx** renameRegEx);

protected:
	CSmartRenameRegEx();
	virtual ~CSmartRenameRegEx();

	void _OnSearchTermChanged();
	void _OnReplaceTermChanged();
	void _OnFlagsChanged();

	size_t _Find(std::wstring data, std::wstring toSearch, bool caseInsensitive, size_t pos);

	uint32_t m_flags = DEFAULT_FLAGS;
	PWSTR m_searchTerm = nullptr;
	PWSTR m_replaceTerm = nullptr;

	CSRWLock m_lock;
	CSRWLock m_lockEvents;

	uint32_t m_cookie = 0;

	struct RENAME_REGEX_EVENT
	{
		ISmartRenameRegExEvents* pEvents;
		uint32_t cookie;
	};

	_Guarded_by_(m_lockEvents) std::vector<RENAME_REGEX_EVENT> m_renameRegExEvents;

	std::atomic<ULONG> m_refCount = 1;
};
