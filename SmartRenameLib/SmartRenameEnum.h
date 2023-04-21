#pragma once

//#include "smartrename_pch.h"
//#include "smartrenameinterfaces.h"
//#include <vector>
//#include "srwlock.h"

class CSmartRenameEnum : public ISmartRenameEnum
{
public:

	// IUnknown
	IFACEMETHODIMP QueryInterface(_In_ REFIID riid,_COM_Outptr_ void** ppv)
	{
		static const QITAB qit[] 
		{
			QITABENT(CSmartRenameEnum, ISmartRenameEnum),
			{0}
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

	// ISmartRenameEnum
	IFACEMETHODIMP Start();
	IFACEMETHODIMP Cancel();

public:
	static HRESULT s_CreateInstance(_In_ IUnknown* punk, _In_ ISmartRenameManager* psrm, _In_ REFIID iid, _Outptr_ void** resultInterface);

protected:
	CSmartRenameEnum(void) noexcept = default;
	virtual ~CSmartRenameEnum(void) noexcept = default;

	CSmartRenameEnum(CSmartRenameEnum const&) = delete;
	CSmartRenameEnum& operator=(CSmartRenameEnum const&) = delete;

	HRESULT _Init(_In_ IUnknown* punk, _In_ ISmartRenameManager* psrm);
	HRESULT _ParseEnumItems(_In_ IEnumShellItems* pesi, _In_ int depth = 0);

	ATL::CComPtr<ISmartRenameManager> m_spsrm;
	ATL::CComPtr<IUnknown> m_spunk;
	std::atomic<ULONG> m_refCount = 1;
	bool m_canceled = false;
	bool _padding[3];
};
