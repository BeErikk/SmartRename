#pragma once

//#include "smartrenameinterfaces.h"

class CMockSmartRenameManagerEvents : public ISmartRenameManagerEvents
{
public:
	CMockSmartRenameManagerEvents()
		: m_refCount(1)
	{
	}

	// IUnknown
	IFACEMETHODIMP QueryInterface(__in REFIID riid, __deref_out void** ppv);
	IFACEMETHODIMP_(ULONG)
	AddRef();
	IFACEMETHODIMP_(ULONG)
	Release();

	// ISmartRenameManagerEvents
	IFACEMETHODIMP OnItemAdded(_In_ ISmartRenameItem* renameItem);
	IFACEMETHODIMP OnUpdate(_In_ ISmartRenameItem* renameItem);
	IFACEMETHODIMP OnError(_In_ ISmartRenameItem* renameItem);
	IFACEMETHODIMP OnRegExStarted(_In_ uint32_t threadId);
	IFACEMETHODIMP OnRegExCanceled(_In_ uint32_t threadId);
	IFACEMETHODIMP OnRegExCompleted(_In_ uint32_t threadId);
	IFACEMETHODIMP OnRenameStarted();
	IFACEMETHODIMP OnRenameCompleted();

	static HRESULT s_CreateInstance(_In_ ISmartRenameManager* psrm, _Outptr_ ISmartRenameUI** ppsrui);

	~CMockSmartRenameManagerEvents()
	{
	}

	ATL::CComPtr<ISmartRenameItem> m_itemAdded;
	ATL::CComPtr<ISmartRenameItem> m_itemUpdated;
	ATL::CComPtr<ISmartRenameItem> m_itemError;
	bool m_regExStarted = false;
	bool m_regExCanceled = false;
	bool m_regExCompleted = false;
	bool m_renameStarted = false;
	bool m_renameCompleted = false;
	std::atomic<ULONG> m_refCount = 1;
};
